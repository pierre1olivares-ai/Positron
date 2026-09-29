package com.timematters.qstar.infrastructure.sharepoint;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicLong;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;
import org.springframework.http.HttpStatus;
import org.springframework.web.client.HttpClientErrorException;

class IssueRepositoryTest {
    private final SharePointRestClient client = mock(SharePointRestClient.class);
    private final SharePointProperties properties = new SharePointProperties();
    private final String root = "/sites/q/Lists/Progress Log";
    private IssueRepository repository;

    @BeforeEach
    void setup() {
        properties.setSiteUrl("https://example.sharepoint.com/sites/q");
        when(client.getItems(eq(properties.getConfigListName()), anyString()))
                .thenReturn(List.of(Map.of("Id", 1, "ReferenceOffset", 1000.0)));
        when(client.listInfo(properties.getProgressListName()))
                .thenReturn(
                        new SharePointRestClient.ListInfo(
                                "progress", root, "SP.Data.ProgressListItem"));
        when(client.getItems(eq(properties.getProgressListName()), anyString()))
                .thenReturn(List.of());
        repository = new IssueRepository(client, properties);
    }

    private Map<String, Object> issue(long id) {
        return new HashMap<>(
                Map.of(
                        "Id",
                        id,
                        "QsNumber",
                        1000.0 + id,
                        "Triaged",
                        "Yes",
                        "ReportDate",
                        "2026-09-28T00:00:00Z",
                        "__metadata",
                        Map.of("etag", "\"1\"")));
    }

    @Test
    void nativePersonReadPreservesIdsEmailCalendarAndEtag() {
        var row = issue(42);
        row.put(
                "TaskOwner",
                Map.of("Id", 7, "Title", "Owner", "Name", "i:0#.f|membership|owner@example.com"));
        row.put("ClosedAt", "2026-09-28T12:13:14Z");
        when(client.getItem(eq(properties.getIssuesListName()), eq(42L), anyString()))
                .thenReturn(row);
        var result = repository.findById(42);
        assertEquals(7L, result.getTaskOwnerId());
        assertEquals("owner@example.com", result.getTaskOwnerEmail());
        assertEquals("2026-09-28", result.getReportDate());
        assertEquals("2026-09-28T12:13:14Z", result.getClosedAt());
        assertEquals("\"1\"", result.getETag());
        assertEquals(1042L, result.getQsNumber());
        assertTrue(result.getTriaged());
    }

    @Test
    void changedEmailOverridesStaleLookupAndChoiceDatesSerializeCorrectly() {
        when(client.ensureUser("new@example.com"))
                .thenReturn(new SharePointUser(8, "New Owner", "new@example.com"));
        var patch = new HashMap<String, Object>();
        patch.put("taskOwnerId", 7L);
        patch.put("taskOwnerEmail", "new@example.com");
        patch.put("triaged", true);
        patch.put("dueDate", null);
        patch.put("qsNumber", 99);
        repository.update(42, patch, "\"3\"");
        ArgumentCaptor<Map<String, Object>> fields = ArgumentCaptor.forClass(Map.class);
        verify(client)
                .updateItem(
                        eq(properties.getIssuesListName()), eq(42L), fields.capture(), eq("\"3\""));
        assertEquals(8L, fields.getValue().get("TaskOwnerId"));
        assertEquals("Yes", fields.getValue().get("Triaged"));
        assertTrue(fields.getValue().containsKey("DueDate"));
        assertNull(fields.getValue().get("DueDate"));
        assertFalse(fields.getValue().containsKey("QsNumber"));
        assertFalse(fields.getValue().containsKey("FollowUp"));
    }

    @Test
    void existingReferencesSurviveAndMissingReferencesUseIdOffset() {
        var legacy = issue(1);
        legacy.put("QsNumber", 27.0);
        var pending = issue(2);
        pending.remove("QsNumber");
        when(client.getItems(eq(properties.getIssuesListName()), anyString()))
                .thenReturn(List.of(legacy, pending));
        var issues = repository.findAll();
        assertEquals(27L, issues.get(0).getQsNumber());
        assertEquals(1002L, issues.get(1).getQsNumber());
    }

    @Test
    void concurrentCreatesUseDistinctServerIdsRatherThanClientMaxima() {
        AtomicLong sequence = new AtomicLong(20);
        var rows = new ConcurrentHashMap<Long, Map<String, Object>>();
        when(client.createItem(eq(properties.getIssuesListName()), anyMap()))
                .thenAnswer(
                        call -> {
                            long id = sequence.incrementAndGet();
                            var row = issue(id);
                            row.remove("QsNumber");
                            rows.put(id, row);
                            return row;
                        });
        when(client.getItem(eq(properties.getIssuesListName()), anyLong(), anyString()))
                .thenAnswer(call -> rows.get(call.getArgument(1)));
        doAnswer(
                        call -> {
                            rows.get(call.<Long>getArgument(1)).putAll(call.getArgument(2));
                            return null;
                        })
                .when(client)
                .updateItem(eq(properties.getIssuesListName()), anyLong(), anyMap(), anyString());
        var input =
                Map.<String, Object>of(
                        "shortSummary",
                        "Issue",
                        "description",
                        "Description",
                        "severity",
                        "Medium");
        var first = CompletableFuture.supplyAsync(() -> repository.create(input));
        var second = CompletableFuture.supplyAsync(() -> repository.create(input));
        var a = first.join();
        var b = second.join();
        assertNotEquals(a.getId(), b.getId());
        assertNotEquals(a.getQsNumber(), b.getQsNumber());
        assertEquals(1000 + a.getId(), a.getQsNumber());
        assertEquals(1000 + b.getId(), b.getQsNumber());
        verify(client, times(2)).createItem(anyString(), anyMap());
    }

    @Test
    void acceptedCreateFailuresKeepIdentityAndNeverPostAgain() {
        when(client.createItem(anyString(), anyMap())).thenReturn(issue(42));
        doThrow(new IllegalStateException("patch unavailable"))
                .when(client)
                .updateItem(anyString(), anyLong(), anyMap(), anyString());
        when(client.getItem(eq(properties.getIssuesListName()), eq(42L), anyString()))
                .thenThrow(new IllegalStateException("read unavailable"));
        var saved =
                repository.create(
                        Map.of(
                                "shortSummary",
                                "Accepted",
                                "description",
                                "Description",
                                "severity",
                                "Medium"));
        assertEquals(42L, saved.getId());
        assertEquals(1042L, saved.getQsNumber());
        assertNull(saved.getETag());
        assertNotNull(saved.getSaveWarning());
        verify(client, times(1)).createItem(anyString(), anyMap());
    }

    @Test
    void journalAssociationUsesExactFolderEvenIfParentMetadataIsForged() {
        var valid =
                new HashMap<String, Object>(
                        Map.of(
                                "Id",
                                9,
                                "FileDirRef",
                                root + "/issue-42",
                                "FSObjType",
                                0,
                                "ParentItemId",
                                99,
                                "EntryText",
                                "Actual issue 42",
                                "Created",
                                "2026-09-28T12:00:00Z",
                                "Author",
                                Map.of("Id", 7, "Title", "Owner", "EMail", "owner@example.com")));
        var loose = new HashMap<>(valid);
        loose.put("Id", 10);
        loose.put("FileDirRef", root);
        loose.put("ParentItemId", 42);
        when(client.getItems(eq(properties.getProgressListName()), anyString()))
                .thenReturn(List.of(valid, loose));
        when(client.getItems(eq(properties.getIssuesListName()), anyString()))
                .thenReturn(List.of(issue(42), issue(99)));
        var all = repository.findAll();
        assertEquals(1, all.get(0).getProgressLog().size());
        assertEquals(9L, all.get(0).getProgressLog().getFirst().getId());
        assertTrue(all.get(1).getProgressLog().isEmpty());
        when(client.getItem(eq(properties.getProgressListName()), eq(9L), anyString()))
                .thenReturn(valid);
        assertThrows(
                com.timematters.qstar.service.IssueOperationException.class,
                () -> repository.progressEntry(99, 9));
    }

    @Test
    void acceptedAppendReadbackFailureKeepsIdAndDoesNotInventCreatedTime() {
        when(client.appendInFolder(properties.getProgressListName(), root + "/issue-42", "Saved"))
                .thenReturn(19L);
        when(client.getItem(eq(properties.getProgressListName()), eq(19L), anyString()))
                .thenThrow(new IllegalStateException("read unavailable"));
        var result =
                repository.append(42, "Saved", new SharePointUser(7, "Owner", "owner@example.com"));
        assertEquals(19L, result.getId());
        assertEquals(7L, result.getAuthorId());
        assertEquals("", result.getTs());
        assertEquals(
                "Your update was posted. Reload its SharePoint-recorded time and details; do not post it again.",
                result.getSaveWarning());
        verify(client, times(1)).appendInFolder(anyString(), anyString(), anyString());
    }

    @Test
    void missingFolderPermissionNeverFallsBackToRootAppend() {
        doThrow(new HttpClientErrorException(HttpStatus.FORBIDDEN))
                .when(client)
                .ensureFolder(anyString());
        var error =
                assertThrows(
                        com.timematters.qstar.service.IssueOperationException.class,
                        () ->
                                repository.append(
                                        42,
                                        "Draft",
                                        new SharePointUser(7, "Owner", "owner@example.com")));
        assertEquals("PROGRESS_ACCESS_PENDING", error.code());
        verify(client, never()).appendInFolder(anyString(), anyString(), anyString());
    }

    @Test
    void historyPreservesNativeFieldsAndDetectsConcurrentRevisionChanges() {
        var current = issue(42);
        current.put("Created", "2025-12-01T00:00:00Z");
        current.put("Modified", "2026-01-05T00:00:00Z");
        current.put("OData__UIVersionString", "3.0");
        var versions =
                List.<Map<String, Object>>of(
                        Map.of(
                                "VersionId",
                                1536,
                                "VersionLabel",
                                "3.0",
                                "Created",
                                "2026-01-05T00:00:00Z",
                                "IsCurrentVersion",
                                true,
                                "Status",
                                "In Progress"));
        when(client.getItem(eq(properties.getIssuesListName()), eq(42L), anyString()))
                .thenReturn(current);
        when(client.getItemVersions(properties.getIssuesListName(), 42)).thenReturn(versions);
        var result = repository.history(42);
        assertEquals(true, result.get("complete"));
        assertEquals("\"1\"", result.get("eTag"));
        assertEquals(current, result.get("current"));
        assertEquals(versions, result.get("versions"));
        assertFalse(
                ((Map<?, ?>) ((List<?>) result.get("versions")).getFirst()).containsKey("Triaged"));
        var newer = new HashMap<>(current);
        newer.put("__metadata", Map.of("etag", "\"2\""));
        when(client.getItem(eq(properties.getIssuesListName()), eq(42L), anyString()))
                .thenReturn(current, newer);
        assertEquals(false, repository.history(42).get("complete"));
        when(client.getItemVersions(properties.getIssuesListName(), 42))
                .thenThrow(new IllegalStateException("History unavailable"));
        assertThrows(IllegalStateException.class, () -> repository.history(42));
        verify(client, never()).updateItem(anyString(), anyLong(), anyMap(), anyString());
    }
}

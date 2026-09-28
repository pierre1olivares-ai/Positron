package com.timematters.qstar.service;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.timematters.qstar.api.model.IssueATO;
import com.timematters.qstar.api.model.IssuePatchATO;
import com.timematters.qstar.api.model.ProgressCreateATO;
import com.timematters.qstar.configuration.security.CallerContext;
import com.timematters.qstar.configuration.security.CurrentUserProvider;
import com.timematters.qstar.infrastructure.sharepoint.IssueRepository;
import com.timematters.qstar.model.Issue;
import java.time.Clock;
import java.time.Instant;
import java.time.ZoneId;
import java.util.Map;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;
import org.openapitools.jackson.nullable.JsonNullableModule;
import org.springframework.http.HttpStatus;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.web.client.HttpClientErrorException;

class IssueServiceTest {
    private final ObjectMapper json = new ObjectMapper().registerModule(new JsonNullableModule());
    private final IssueRepository repository = mock(IssueRepository.class);
    private final CurrentUserProvider users = mock(CurrentUserProvider.class);
    private final Clock clock =
            Clock.fixed(Instant.parse("2026-09-28T12:00:00Z"), ZoneId.of("Europe/Amsterdam"));
    private IssueService service;

    @BeforeEach
    void setup() {
        when(users.currentUser())
                .thenReturn(new CallerContext("qm", 7, "Quality", "qm@example.com"));
        service = new IssueService(repository, users, json, clock);
    }

    private Issue issue() {
        Issue issue = new Issue();
        issue.setId(42L);
        issue.setQsNumber(1042L);
        issue.setETag("\"1\"");
        issue.setStatus("In Progress");
        issue.setTransformedInto("OFI");
        issue.setTaskOwnerId(7L);
        issue.setReportDate("2026-09-01");
        issue.setSeverity("Medium");
        issue.setTriaged(true);
        issue.setShortSummary("Issue");
        issue.setDescription("Description");
        issue.setTaskCreated("Yes");
        return issue;
    }

    @Test
    void absentAndExplicitNullRemainDifferentThroughRealJsonBinding() throws Exception {
        Issue current = issue();
        when(repository.findById(42)).thenReturn(current);
        IssuePatchATO patch =
                json.readValue("{\"dueDate\":null,\"followUp\":\"changed\"}", IssuePatchATO.class);
        assertInstanceOf(IssueATO.class, service.updateIssue(42, patch, "\"1\""));
        ArgumentCaptor<Map<String, Object>> fields = ArgumentCaptor.forClass(Map.class);
        verify(repository).update(eq(42L), fields.capture(), eq("\"1\""));
        assertTrue(fields.getValue().containsKey("dueDate"));
        assertEquals("", fields.getValue().get("dueDate"));
        assertEquals("changed", fields.getValue().get("followUp"));
        assertFalse(fields.getValue().containsKey("taskOwner"));
    }

    @Test
    void staleAndWildcardVersionsNeverWrite() throws Exception {
        Issue latest = issue();
        latest.setETag("\"2\"");
        when(repository.findById(42)).thenReturn(latest);
        IssuePatchATO patch = json.readValue("{\"followUp\":\"my draft\"}", IssuePatchATO.class);
        IssueOperationException stale =
                assertThrows(
                        IssueOperationException.class,
                        () -> service.updateIssue(42, patch, "\"1\""));
        assertEquals(412, stale.status());
        assertEquals("ISSUE_CONFLICT", stale.code());
        assertEquals(
                428,
                assertThrows(
                                IssueOperationException.class,
                                () -> service.updateIssue(42, patch, "*"))
                        .status());
        verify(repository, never()).update(anyLong(), anyMap(), anyString());
    }

    @Test
    void serverConflictAfterPreflightRemainsAConflict() throws Exception {
        Issue current = issue();
        when(repository.findById(42)).thenReturn(current);
        doThrow(new HttpClientErrorException(HttpStatus.PRECONDITION_FAILED))
                .when(repository)
                .update(anyLong(), anyMap(), anyString());
        IssuePatchATO patch = json.readValue("{\"followUp\":\"my draft\"}", IssuePatchATO.class);
        assertEquals(
                412,
                assertThrows(
                                IssueOperationException.class,
                                () -> service.updateIssue(42, patch, "\"1\""))
                        .status());
        verify(repository).update(eq(42L), anyMap(), eq("\"1\""));
    }

    @Test
    void acceptedMutationWithFailedReadbackReturnsSavedReceipt() throws Exception {
        when(repository.findById(42))
                .thenReturn(issue())
                .thenThrow(new IllegalStateException("read unavailable"));
        Object result =
                service.updateIssue(
                        42,
                        json.readValue("{\"followUp\":\"saved\"}", IssuePatchATO.class),
                        "\"1\"");
        assertInstanceOf(Map.class, result);
        assertEquals(true, ((Map<?, ?>) result).get("saved"));
        assertEquals(42L, ((Map<?, ?>) result).get("issueId"));
        verify(repository, times(1)).update(eq(42L), anyMap(), anyString());
    }

    @Test
    void ownerCannotReassignOrMutateAnotherOwnersIssue() throws Exception {
        when(users.currentUser())
                .thenReturn(new CallerContext("owner", 7, "Owner", "owner@example.com"));
        Issue current = issue();
        when(repository.findById(42)).thenReturn(current);
        IssuePatchATO reassign = json.readValue("{\"taskOwnerId\":8}", IssuePatchATO.class);
        assertThrows(AccessDeniedException.class, () -> service.updateIssue(42, reassign, "\"1\""));
        current.setTaskOwnerId(8L);
        IssuePatchATO status = json.readValue("{\"status\":\"Created\"}", IssuePatchATO.class);
        assertThrows(AccessDeniedException.class, () -> service.updateIssue(42, status, "\"1\""));
        verify(repository, never()).update(anyLong(), anyMap(), anyString());
    }

    @Test
    void ownerStatusUpdateDerivesAuditFieldsInsteadOfTrustingRequest() throws Exception {
        when(users.currentUser())
                .thenReturn(new CallerContext("owner", 7, "Owner", "owner@example.com"));
        when(repository.findById(42)).thenReturn(issue());
        IssuePatchATO patch =
                json.readValue(
                        "{\"status\":\"Created\",\"ownerUpdateAt\":\"forged\",\"ownerUpdateText\":\"forged\"}",
                        IssuePatchATO.class);
        service.updateIssue(42, patch, "\"1\"");
        ArgumentCaptor<Map<String, Object>> fields = ArgumentCaptor.forClass(Map.class);
        verify(repository).update(eq(42L), fields.capture(), anyString());
        assertEquals(clock.instant().toString(), fields.getValue().get("ownerUpdateAt"));
        assertEquals(true, fields.getValue().get("ownerUpdate"));
        assertNotEquals("forged", fields.getValue().get("ownerUpdateText"));
    }

    @Test
    void closureUsesTheSameNcGateAndOfiExemption() throws Exception {
        Issue current = issue();
        current.setTransformedInto("NC Minor");
        when(repository.findById(42)).thenReturn(current);
        IssuePatchATO close = json.readValue("{\"status\":\"Closed\"}", IssuePatchATO.class);
        assertThrows(IssueOperationException.class, () -> service.updateIssue(42, close, "\"1\""));
        current.setTransformedInto("OFI");
        service.updateIssue(42, close, "\"1\"");
        ArgumentCaptor<Map<String, Object>> fields = ArgumentCaptor.forClass(Map.class);
        verify(repository).update(eq(42L), fields.capture(), anyString());
        assertEquals("2026-09-28", fields.getValue().get("closedDate"));
    }

    @Test
    void reopenedIssueClearsOldVerificationAndStartsANewReminderCycle() throws Exception {
        Issue current = issue();
        current.setStatus("Closed");
        when(repository.findById(42)).thenReturn(current);
        service.updateIssue(
                42, json.readValue("{\"status\":\"In Progress\"}", IssuePatchATO.class), "\"1\"");
        ArgumentCaptor<Map<String, Object>> fields = ArgumentCaptor.forClass(Map.class);
        verify(repository).update(eq(42L), fields.capture(), anyString());
        assertEquals("", fields.getValue().get("closedDate"));
        assertNull(fields.getValue().get("verifiedById"));
        assertEquals(clock.instant().toString(), fields.getValue().get("reminderCycle"));
    }

    @Test
    void triageNormalizesReportEnvelopeAndWritesBooleanDomainValue() throws Exception {
        Issue current = issue();
        current.setTriaged(false);
        current.setDueDate("");
        current.setReportDate("2026-09-01T00:00:00Z");
        when(repository.findById(42)).thenReturn(current);
        service.updateIssue(42, json.readValue("{\"triaged\":true}", IssuePatchATO.class), "\"1\"");
        ArgumentCaptor<Map<String, Object>> fields = ArgumentCaptor.forClass(Map.class);
        verify(repository).update(eq(42L), fields.capture(), anyString());
        assertEquals("2026-10-01", fields.getValue().get("dueDate"));
        assertEquals(true, fields.getValue().get("triaged"));
    }

    @Test
    void rejectedOrClosedIssueCannotReceiveProgress() {
        Issue current = issue();
        current.setStatus("Closed");
        when(repository.findById(42)).thenReturn(current);
        ProgressCreateATO input = new ProgressCreateATO();
        input.setText("Comment");
        assertThrows(IssueOperationException.class, () -> service.addProgressLogEntry(42, input));
        verify(repository, never()).append(anyLong(), anyString(), any());
    }

    @Test
    void calendarDateEnvelopesAndMonthEndAreStable() {
        assertEquals("2026-01-01", CalendarDates.normalize("2026-01-01T00:00:00+14:00"));
        assertEquals("2026-01-01", CalendarDates.normalize("2026-01-01T00:00:00-12:00"));
        Issue current = issue();
        current.setStatus(IssueLifecycle.TEST);
        current.setTransformedInto("NC Major");
        current.setImplementationDate("2026-07-31");
        current.setVerifiedBy("Verifier");
        current.setVerifiedById(8L);
        var patch = new java.util.HashMap<String, Object>();
        patch.put("status", "Closed");
        assertThrows(
                IssueOperationException.class, () -> IssueLifecycle.apply(current, patch, clock));
        IssueLifecycle.apply(
                current,
                patch,
                Clock.fixed(Instant.parse("2026-09-30T12:00:00Z"), ZoneId.of("Europe/Amsterdam")));
        assertEquals("2026-09-30", patch.get("closedDate"));
    }
}

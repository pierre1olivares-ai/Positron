package com.timematters.qstar.api;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.ArgumentMatchers.*;
import static org.mockito.Mockito.*;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.*;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.timematters.qstar.api.controller.IssuesApiController;
import com.timematters.qstar.configuration.security.AuthorizationPolicy;
import com.timematters.qstar.configuration.security.CallerContext;
import com.timematters.qstar.configuration.security.CurrentUserProvider;
import com.timematters.qstar.infrastructure.sharepoint.IssueRepository;
import com.timematters.qstar.model.Issue;
import com.timematters.qstar.model.ProgressLogEntry;
import com.timematters.qstar.service.IssueService;
import java.time.Clock;
import java.time.Instant;
import java.time.ZoneOffset;
import java.util.Map;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;
import org.openapitools.jackson.nullable.JsonNullableModule;
import org.springframework.http.MediaType;
import org.springframework.http.converter.json.MappingJackson2HttpMessageConverter;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.setup.MockMvcBuilders;
import org.springframework.web.client.RestClientException;

class IssuesHttpContractTest {
    private final IssueRepository repository = mock(IssueRepository.class);
    private final CurrentUserProvider users = mock(CurrentUserProvider.class);
    private final ObjectMapper json = new ObjectMapper().registerModule(new JsonNullableModule());
    private MockMvc mvc;

    @BeforeEach
    void setup() {
        when(users.currentUser())
                .thenReturn(new CallerContext("qm", 7, "Quality", "qm@example.com"));
        IssueService service =
                new IssueService(
                        repository,
                        users,
                        json,
                        Clock.fixed(Instant.parse("2026-09-28T12:00:00Z"), ZoneOffset.UTC));
        mvc =
                MockMvcBuilders.standaloneSetup(
                                new IssuesApiController(
                                        service, new AuthorizationPolicy(users), "/api/v1"))
                        .addPlaceholderValue("openapi.qStarIssueManager.base-path", "/api/v1")
                        .setControllerAdvice(new CustomExceptionHandler())
                        .setMessageConverters(new MappingJackson2HttpMessageConverter(json))
                        .build();
    }

    private Issue issue() {
        Issue issue = new Issue();
        issue.setId(42L);
        issue.setQsNumber(1042L);
        issue.setETag("\"3\"");
        issue.setTriaged(true);
        issue.setTaskCreated("Yes");
        issue.setShortSummary("Issue");
        issue.setDescription("Description");
        issue.setSeverity("Medium");
        issue.setStatus("In Progress");
        issue.setTransformedInto("OFI");
        issue.setTaskOwner("Owner");
        issue.setTaskOwnerId(8L);
        issue.setTaskOwnerEmail("owner@example.com");
        issue.setReportDate("2026-09-28");
        issue.setReminderCycle("initial");
        return issue;
    }

    @Test
    void actualGeneratedDtoSerializationPreservesEtagAndNativePeople() throws Exception {
        when(repository.findById(42)).thenReturn(issue());
        mvc.perform(get("/api/v1/issues/42"))
                .andExpect(status().isOk())
                .andExpect(header().string("ETag", "\"3\""))
                .andExpect(jsonPath("$.eTag").value("\"3\""))
                .andExpect(jsonPath("$.taskOwnerId").value(8))
                .andExpect(jsonPath("$.taskOwnerEmail").value("owner@example.com"))
                .andExpect(jsonPath("$.reportDate").value("2026-09-28"))
                .andExpect(jsonPath("$.reminderCycle").value("initial"));
    }

    @Test
    void httpPatchPreservesNullPresenceAndOriginalHeaderThroughDomainService() throws Exception {
        when(repository.findById(42)).thenReturn(issue());
        mvc.perform(
                        patch("/api/v1/issues/42")
                                .header("If-Match", "\"3\"")
                                .contentType(MediaType.APPLICATION_JSON)
                                .content("{\"dueDate\":null,\"followUp\":\"saved\"}"))
                .andExpect(status().isOk());
        ArgumentCaptor<Map<String, Object>> fields = ArgumentCaptor.forClass(Map.class);
        verify(repository).update(eq(42L), fields.capture(), eq("\"3\""));
        assertEquals("", fields.getValue().get("dueDate"));
        assertTrue(fields.getValue().containsKey("dueDate"));
        assertFalse(fields.getValue().containsKey("taskOwnerId"));
        assertFalse(fields.getValue().containsKey("status"));
    }

    @Test
    void nullEmptyAndInvalidCalendarInputsReturn400WithoutWriting() throws Exception {
        when(repository.findById(42)).thenReturn(issue());
        for (String body :
                new String[] {
                    "{\"status\":null}", "{\"status\":\"\"}", "{\"dueDate\":\"2026-02-31\"}"
                }) {
            mvc.perform(
                            patch("/api/v1/issues/42")
                                    .header("If-Match", "\"3\"")
                                    .contentType(MediaType.APPLICATION_JSON)
                                    .content(body))
                    .andExpect(status().isBadRequest());
        }
        verify(repository, never()).update(anyLong(), anyMap(), anyString());
    }

    @Test
    void acceptedWriteReceiptIs200AndTypedConflictIs412() throws Exception {
        when(repository.findById(42))
                .thenReturn(issue())
                .thenThrow(new IllegalStateException("read failed"));
        mvc.perform(
                        patch("/api/v1/issues/42")
                                .header("If-Match", "\"3\"")
                                .contentType(MediaType.APPLICATION_JSON)
                                .content("{\"followUp\":\"saved\"}"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.saved").value(true))
                .andExpect(jsonPath("$.issueId").value(42));
        doReturn(issue()).when(repository).findById(42);
        mvc.perform(
                        patch("/api/v1/issues/42")
                                .header("If-Match", "\"2\"")
                                .contentType(MediaType.APPLICATION_JSON)
                                .content("{\"followUp\":\"draft\"}"))
                .andExpect(status().isPreconditionFailed())
                .andExpect(jsonPath("$.code").value("ISSUE_CONFLICT"))
                .andExpect(jsonPath("$.meta.freshIssue.eTag").value("\"3\""));
    }

    @Test
    void acceptedIdentityLocationsResolveToRegisteredReadEndpoints() throws Exception {
        when(repository.create(anyMap())).thenReturn(issue());
        when(repository.findById(42)).thenReturn(issue());
        String location =
                mvc.perform(
                                post("/api/v1/issues")
                                        .contentType(MediaType.APPLICATION_JSON)
                                        .content(
                                                "{\"shortSummary\":\"Issue\",\"description\":\"Description\",\"severity\":\"Medium\"}"))
                        .andExpect(status().isCreated())
                        .andExpect(header().string("X-QStar-Reference", "1042"))
                        .andExpect(header().string("Location", "/api/v1/issues/42"))
                        .andReturn()
                        .getResponse()
                        .getHeader("Location");
        mvc.perform(get(location)).andExpect(status().isOk());
        ProgressLogEntry entry = new ProgressLogEntry();
        entry.setId(19L);
        entry.setText("Posted");
        entry.setAuthor("Quality");
        entry.setAuthorId(7L);
        entry.setAuthorEmail("qm@example.com");
        entry.setTs("");
        entry.setSaveWarning("Posted; reload audit time");
        when(repository.append(eq(42L), eq("Posted"), any())).thenReturn(entry);
        when(repository.progressEntry(42, 19)).thenReturn(entry);
        String journal =
                mvc.perform(
                                post("/api/v1/issues/42/progress")
                                        .contentType(MediaType.APPLICATION_JSON)
                                        .content("{\"text\":\"Posted\"}"))
                        .andExpect(status().isCreated())
                        .andExpect(header().string("X-QStar-Entry-Id", "19"))
                        .andExpect(jsonPath("$.authorId").value(7))
                        .andExpect(jsonPath("$.ts").value(""))
                        .andReturn()
                        .getResponse()
                        .getHeader("Location");
        assertEquals("/api/v1/issues/42/progress/19", journal);
        mvc.perform(get(journal)).andExpect(status().isOk());
    }

    @Test
    void unexpectedTransportStatusReturns502WithoutAcceptedReceiptOrReadback() throws Exception {
        for (int status : new int[] {302, 307}) {
            for (String operation : new String[] {"create", "update", "progress"}) {
                reset(repository);
                when(repository.findById(42)).thenReturn(issue());
                var failure = new RestClientException("HTTP " + status + " private transport data");
                var request =
                        switch (operation) {
                            case "create" -> {
                                when(repository.create(anyMap())).thenThrow(failure);
                                yield post("/api/v1/issues")
                                        .content(
                                                "{\"shortSummary\":\"Draft\",\"description\":\"Description\",\"severity\":\"Medium\"}");
                            }
                            case "update" -> {
                                doThrow(failure)
                                        .when(repository)
                                        .update(eq(42L), anyMap(), eq("\"3\""));
                                yield patch("/api/v1/issues/42")
                                        .header("If-Match", "\"3\"")
                                        .content("{\"followUp\":\"Draft\"}");
                            }
                            default -> {
                                when(repository.append(eq(42L), eq("Draft"), any()))
                                        .thenThrow(failure);
                                yield post("/api/v1/issues/42/progress")
                                        .content("{\"text\":\"Draft\"}");
                            }
                        };
                var response =
                        mvc.perform(request.contentType(MediaType.APPLICATION_JSON))
                                .andExpect(status().isBadGateway())
                                .andExpect(jsonPath("$.code").value("SHAREPOINT_UNAVAILABLE"))
                                .andExpect(jsonPath("$.saved").doesNotExist())
                                .andExpect(header().doesNotExist("Location"))
                                .andReturn()
                                .getResponse();
                assertFalse(response.getContentAsString().contains("private transport data"));
                if (operation.equals("create")) verify(repository).create(anyMap());
                else {
                    verify(repository).findById(42);
                    if (operation.equals("update"))
                        verify(repository).update(eq(42L), anyMap(), eq("\"3\""));
                    else verify(repository).append(eq(42L), eq("Draft"), any());
                }
                verifyNoMoreInteractions(repository);
            }
        }
    }

    @Test
    void historyGetPreservesRawCoverageAndRequiresNoMutation() throws Exception {
        when(users.currentUser())
                .thenReturn(new CallerContext("reader", 9, "Reader", "reader@example.com"));
        when(repository.history(42))
                .thenReturn(
                        Map.of(
                                "issueId",
                                42L,
                                "eTag",
                                "\"3\"",
                                "complete",
                                false,
                                "current",
                                Map.of("Id", 42L, "OData__UIVersionString", "3.0"),
                                "versions",
                                java.util.List.of(
                                        Map.of(
                                                "VersionId",
                                                1536,
                                                "VersionLabel",
                                                "3.0",
                                                "Triaged",
                                                "Yes"))));
        mvc.perform(get("/api/v1/issues/42/history"))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.issueId").value(42))
                .andExpect(jsonPath("$.eTag").value("\"3\""))
                .andExpect(jsonPath("$.complete").value(false))
                .andExpect(jsonPath("$.current.OData__UIVersionString").value("3.0"))
                .andExpect(jsonPath("$.versions[0].VersionId").value(1536))
                .andExpect(jsonPath("$.versions[0].Triaged").value("Yes"))
                .andExpect(jsonPath("$.versions[0].TaskCreated").doesNotExist());
        verify(repository, never()).update(anyLong(), anyMap(), anyString());
        verify(repository, never()).append(anyLong(), anyString(), any());
    }
}

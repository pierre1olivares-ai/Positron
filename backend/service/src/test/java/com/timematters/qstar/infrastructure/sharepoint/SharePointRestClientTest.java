package com.timematters.qstar.infrastructure.sharepoint;

import static org.junit.jupiter.api.Assertions.*;
import static org.springframework.test.web.client.match.MockRestRequestMatchers.*;
import static org.springframework.test.web.client.response.MockRestResponseCreators.*;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.timematters.qstar.configuration.security.SharePointAccessTokenProvider;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.mock.http.client.MockClientHttpRequest;
import org.springframework.test.web.client.MockRestServiceServer;
import org.springframework.web.client.HttpClientErrorException;
import org.springframework.web.client.RestClientException;
import org.springframework.web.client.RestTemplate;

import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

class SharePointRestClientTest {
    private final ObjectMapper json = new ObjectMapper();
    private final RestTemplate rest = new RestTemplate();
    private final MockRestServiceServer server = MockRestServiceServer.bindTo(rest).build();
    private final SharePointProperties properties = new SharePointProperties();
    private final AtomicInteger tokens = new AtomicInteger();
    private SharePointRestClient client;
    private static final String BASE = "https://example.sharepoint.com/sites/q/_api/";

    @BeforeEach
    void setup() {
        properties.setSiteUrl("https://example.sharepoint.com/sites/q");
        SharePointAccessTokenProvider provider = () -> "caller-token-" + tokens.incrementAndGet();
        client = new SharePointRestClient(properties, provider, rest);
    }

    private JsonNode body(org.springframework.http.client.ClientHttpRequest request)
            throws java.io.IOException {
        return json.readTree(((MockClientHttpRequest) request).getBodyAsString());
    }

    @Test
    void partialMutationTransmitsOriginalVersionNullAndTypedChoiceWithoutAppIdentity() {
        server.expect(
                        requestTo(
                                BASE
                                        + "web/lists/GetByTitle('Issues')?$select=Id,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder"))
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"Id\":\"list\",\"ListItemEntityTypeFullName\":\"SP.Data.IssuesListItem\",\"RootFolder\":{\"ServerRelativeUrl\":\"/sites/q/Lists/Issues\"}}}",
                                MediaType.APPLICATION_JSON));
        server.expect(requestTo(BASE + "web/lists/GetByTitle('Issues')/items(42)"))
                .andExpect(method(HttpMethod.POST))
                .andExpect(header("X-HTTP-Method", "MERGE"))
                .andExpect(header("IF-MATCH", "\"4\""))
                .andExpect(header("Authorization", "Bearer caller-token-2"))
                .andExpect(
                        request -> {
                            JsonNode fields = body(request);
                            assertEquals("Yes", fields.path("Triaged").asText());
                            assertTrue(fields.get("DueDate").isNull());
                            assertEquals(
                                    "SP.Data.IssuesListItem",
                                    fields.path("__metadata").path("type").asText());
                            assertFalse(fields.has("TaskOwnerId"));
                        })
                .andRespond(withNoContent());
        var fields = new java.util.HashMap<String, Object>();
        fields.put("Triaged", "Yes");
        fields.put("DueDate", null);
        client.updateItem("Issues", 42, fields, "\"4\"");
        server.verify();
    }

    @Test
    void mismatchRemains412AndIsNotRetried() {
        server.expect(anything())
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"Id\":\"list\",\"ListItemEntityTypeFullName\":\"SP.Data.IssuesListItem\",\"RootFolder\":{}}}",
                                MediaType.APPLICATION_JSON));
        server.expect(anything())
                .andExpect(header("IF-MATCH", "\"1\""))
                .andRespond(withStatus(HttpStatus.PRECONDITION_FAILED));
        var error =
                assertThrows(
                        HttpClientErrorException.class,
                        () ->
                                client.updateItem(
                                        "Issues", 42, Map.of("FollowUp", "draft"), "\"1\""));
        assertEquals(412, error.getStatusCode().value());
        server.verify();
    }

    @Test
    void delegatedEnsureUserUsesTheSiteIdentityEndpoint() {
        server.expect(requestTo(BASE + "web/ensureuser"))
                .andExpect(method(HttpMethod.POST))
                .andExpect(
                        request ->
                                assertEquals(
                                        "new@example.com",
                                        body(request).path("logonName").asText()))
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"Id\":8,\"Title\":\"New"
                                    + " owner\",\"Email\":\"new@example.com\"}}",
                                MediaType.APPLICATION_JSON));
        assertEquals(8, client.ensureUser("new@example.com").id());
        server.verify();
    }

    @Test
    void journalUsesAbsoluteFolderAndOnlyWritableContentFields() {
        server.expect(
                        requestTo(
                                BASE
                                        + "web/lists/GetByTitle('Progress')/AddValidateUpdateItemUsingPath"))
                .andExpect(method(HttpMethod.POST))
                .andExpect(
                        request -> {
                            JsonNode payload = body(request);
                            assertEquals(
                                    "https://example.sharepoint.com/sites/q/Lists/Progress"
                                        + " Log/issue-42",
                                    payload.path("listItemCreateInfo")
                                            .path("FolderPath")
                                            .path("DecodedUrl")
                                            .asText());
                            assertEquals(
                                    0,
                                    payload.path("listItemCreateInfo")
                                            .path("UnderlyingObjectType")
                                            .asInt());
                            assertFalse(payload.path("bNewDocumentUpdate").asBoolean());
                            assertEquals(2, payload.path("formValues").size());
                            assertEquals(
                                    "Title",
                                    payload.path("formValues").get(0).path("FieldName").asText());
                            assertEquals(
                                    "EntryText",
                                    payload.path("formValues").get(1).path("FieldName").asText());
                            assertEquals(
                                    "Accepted",
                                    payload.path("formValues").get(1).path("FieldValue").asText());
                        })
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"AddValidateUpdateItemUsingPath\":{\"results\":[{\"FieldName\":\"Id\",\"FieldValue\":\"19\",\"HasException\":false}]}}}",
                                MediaType.APPLICATION_JSON));
        assertEquals(
                19,
                client.appendInFolder(
                        "Progress", "/sites/q/Lists/Progress Log/issue-42", "Accepted"));
        server.verify();
    }

    @Test
    void unreadableAcceptedCreateIsDistinguishedFromRejectedCreate() {
        server.expect(anything())
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"ListItemEntityTypeFullName\":\"SP.Data.IssuesListItem\",\"RootFolder\":{}}}",
                                MediaType.APPLICATION_JSON));
        server.expect(anything()).andRespond(withStatus(HttpStatus.CREATED).body("unreadable"));
        assertThrows(
                AcceptedWriteException.class,
                () -> client.createItem("Issues", Map.of("Title", "accepted")));
        server.verify();
    }

    @Test
    void redirectsRejectCreateUpdateAndAppendWithoutFollowupRequestsOrAcceptedReceipts() {
        for (HttpStatus status :
                new HttpStatus[] {HttpStatus.FOUND, HttpStatus.TEMPORARY_REDIRECT}) {
            for (String operation : new String[] {"create", "update", "progress"}) {
                server.reset();
                if (operation.equals("create"))
                    server.expect(anything())
                            .andExpect(method(HttpMethod.GET))
                            .andRespond(
                                    withSuccess(
                                            "{\"d\":{\"results\":[{\"Id\":1,\"ReferenceOffset\":3000}]}}",
                                            MediaType.APPLICATION_JSON));
                server.expect(anything())
                        .andExpect(method(HttpMethod.GET))
                        .andRespond(
                                withSuccess(
                                        "{\"d\":{\"ListItemEntityTypeFullName\":\"SP.Data.IssuesListItem\",\"RootFolder\":{\"ServerRelativeUrl\":\"/sites/q/Lists/Progress\"}}}",
                                        MediaType.APPLICATION_JSON));
                if (operation.equals("progress"))
                    server.expect(anything())
                            .andExpect(method(HttpMethod.GET))
                            .andRespond(
                                    withSuccess(
                                            "{\"d\":{\"Exists\":true}}",
                                            MediaType.APPLICATION_JSON));
                server.expect(anything())
                        .andExpect(method(HttpMethod.POST))
                        .andRespond(
                                withStatus(status)
                                        .header("Location", "https://outside.invalid/private")
                                        .body("<html>private response with caller-token</html>"));
                var repository = new IssueRepository(client, properties);
                var error =
                        assertThrows(
                                RestClientException.class,
                                () -> {
                                    switch (operation) {
                                        case "create" ->
                                                repository.create(Map.of("shortSummary", "Draft"));
                                        case "update" ->
                                                repository.update(
                                                        42, Map.of("followUp", "Draft"), "\"1\"");
                                        case "progress" ->
                                                repository.append(
                                                        42,
                                                        "Draft",
                                                        new SharePointUser(
                                                                7, "Caller", "caller@example.com"));
                                        default -> fail(operation);
                                    }
                                });
                assertEquals(
                        "SharePoint returned an unexpected HTTP status (" + status.value() + ").",
                        error.getMessage());
                assertNull(error.getCause());
                server.verify();
            }
        }
    }

    @Test
    void malformed201AppendStillReturnsAnAcceptedReceiptWithoutRereading() {
        server.expect(anything())
                .andExpect(method(HttpMethod.GET))
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"RootFolder\":{\"ServerRelativeUrl\":\"/sites/q/Lists/Progress\"}}}",
                                MediaType.APPLICATION_JSON));
        server.expect(anything())
                .andExpect(method(HttpMethod.GET))
                .andRespond(withSuccess("{\"d\":{\"Exists\":true}}", MediaType.APPLICATION_JSON));
        server.expect(anything())
                .andExpect(method(HttpMethod.POST))
                .andRespond(withStatus(HttpStatus.CREATED).body("malformed"));
        var saved =
                new IssueRepository(client, properties)
                        .append(
                                42,
                                "Accepted",
                                new SharePointUser(7, "Caller", "caller@example.com"));
        assertNull(saved.getId());
        assertEquals("Accepted", saved.getText());
        assertEquals("", saved.getTs());
        assertNotNull(saved.getSaveWarning());
        server.verify();
    }

    @Test
    void unusableAcceptedCreateShapesAreQuarantinedAcrossTransportAndRepository() {
        for (String response :
                new String[] {"[]", "123", "{\"d\":null}", "{\"d\":{\"Id\":\"not-an-id\"}}"}) {
            server.reset();
            expectCreate(response, 1000);
            var saved =
                    new IssueRepository(client, properties)
                            .create(Map.of("shortSummary", "Accepted"));
            assertNull(saved.getId(), response);
            assertNotNull(saved.getSaveWarning(), response);
            server.verify();
        }
    }

    @Test
    void acceptedIdentityWithUnsupportedReferenceIsStillNeverRetried() {
        expectCreate("{\"d\":{\"Id\":42}}", 9007199254740991L);
        var saved =
                new IssueRepository(client, properties).create(Map.of("shortSummary", "Accepted"));
        assertEquals(42L, saved.getId());
        assertNull(saved.getQsNumber());
        assertNotNull(saved.getSaveWarning());
        server.verify();
    }

    private void expectCreate(String response, long offset) {
        server.expect(anything())
                .andExpect(method(HttpMethod.GET))
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"results\":[{\"Id\":1,\"ReferenceOffset\":"
                                        + offset
                                        + "}]}}",
                                MediaType.APPLICATION_JSON));
        server.expect(anything())
                .andExpect(method(HttpMethod.GET))
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"ListItemEntityTypeFullName\":\"SP.Data.IssuesListItem\",\"RootFolder\":{}}}",
                                MediaType.APPLICATION_JSON));
        server.expect(anything())
                .andExpect(method(HttpMethod.POST))
                .andRespond(
                        withStatus(HttpStatus.CREATED)
                                .contentType(MediaType.APPLICATION_JSON)
                                .body(response));
    }

    @Test
    void pagingRetainsMetadataAndWillNotSendCallerTokenOutsideConfiguredSite() {
        server.expect(anything())
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"results\":[{\"Id\":1}],\"__next\":\"https://example.sharepoint.com/sites/q/_api/web/lists/GetByTitle('Issues')/items?$skiptoken=Paged%3DTRUE%26p_ID%3D1\"}}",
                                MediaType.APPLICATION_JSON));
        server.expect(anything())
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"results\":[{\"Id\":2,\"__metadata\":{\"etag\":\"\\\"2\\\"\"}}]}}",
                                MediaType.APPLICATION_JSON));
        var rows = client.getItems("Issues", "");
        assertEquals(2, rows.size());
        assertEquals("\"2\"", IssueRepository.eTag(rows.get(1)));
        server.verify();
        server.reset();
        server.expect(anything())
                .andRespond(
                        withSuccess(
                                "{\"d\":{\"results\":[],\"__next\":\"https://evil.invalid/collect\"}}",
                                MediaType.APPLICATION_JSON));
        assertThrows(IllegalArgumentException.class, () -> client.getItems("Issues", ""));
        server.verify();
    }
}

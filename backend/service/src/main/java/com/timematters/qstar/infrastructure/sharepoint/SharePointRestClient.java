package com.timematters.qstar.infrastructure.sharepoint;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import com.timematters.qstar.configuration.security.SharePointAccessTokenProvider;
import java.net.URI;
import java.net.http.HttpClient;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.ResponseEntity;
import org.springframework.http.client.JdkClientHttpRequestFactory;
import org.springframework.stereotype.Component;
import org.springframework.web.client.RestTemplate;
import org.springframework.web.util.UriUtils;

/** All list operations run with the caller's delegated token against one configured site. */
@Component
public class SharePointRestClient {
    private final SharePointProperties properties;
    private final SharePointAccessTokenProvider tokens;
    private final RestTemplate rest;
    private final ObjectMapper json = new ObjectMapper();

    @Autowired
    public SharePointRestClient(
            SharePointProperties properties, SharePointAccessTokenProvider tokens) {
        this(properties, tokens, new RestTemplate(requestFactory()));
    }

    public SharePointRestClient(
            SharePointProperties properties,
            SharePointAccessTokenProvider tokens,
            RestTemplate rest) {
        this.properties = properties;
        this.tokens = tokens;
        this.rest = rest;
    }

    private static JdkClientHttpRequestFactory requestFactory() {
        var http =
                HttpClient.newBuilder()
                        .connectTimeout(Duration.ofSeconds(20))
                        .followRedirects(HttpClient.Redirect.NEVER)
                        .build();
        var factory = new JdkClientHttpRequestFactory(http);
        factory.setReadTimeout(Duration.ofSeconds(40));
        return factory;
    }

    public record ListInfo(String id, String rootPath, String entityType) {}

    public String getSiteId() {
        return request(HttpMethod.GET, "web?$select=Id", null, null, false).path("Id").asText();
    }

    public SharePointUser currentUser() {
        JsonNode user =
                request(
                        HttpMethod.GET,
                        "web/currentuser?$select=Id,Title,Email",
                        null,
                        null,
                        false);
        return new SharePointUser(
                user.path("Id").asLong(),
                user.path("Title").asText(""),
                user.path("Email").asText(""));
    }

    public SharePointUser ensureUser(String email) {
        JsonNode user =
                request(HttpMethod.POST, "web/ensureuser", Map.of("logonName", email), null, false);
        return new SharePointUser(
                user.path("Id").asLong(),
                user.path("Title").asText(""),
                user.path("Email").asText(""));
    }

    public ListInfo listInfo(String title) {
        JsonNode list =
                request(
                        HttpMethod.GET,
                        listPath(title)
                                + "?$select=Id,ListItemEntityTypeFullName,RootFolder/ServerRelativeUrl&$expand=RootFolder",
                        null,
                        null,
                        false);
        return new ListInfo(
                list.path("Id").asText(),
                list.path("RootFolder").path("ServerRelativeUrl").asText(),
                list.path("ListItemEntityTypeFullName").asText());
    }

    public String getListId(String title) {
        return listInfo(title).id();
    }

    public List<Map<String, Object>> getItems(String title, String query) {
        var rows = new ArrayList<Map<String, Object>>();
        String next =
                listPath(title)
                        + "/items"
                        + (query.isEmpty() ? "?$top=2000" : "?" + query + "&$top=2000");
        while (next != null) {
            JsonNode page = request(HttpMethod.GET, next, null, null, false);
            JsonNode values = page.has("results") ? page.get("results") : page.path("value");
            if (!values.isArray())
                throw new IllegalStateException("SharePoint returned an invalid item collection.");
            for (JsonNode row : values) rows.add(toMap(row));
            next =
                    page.has("__next")
                            ? page.get("__next").asText()
                            : page.has("odata.nextLink")
                                    ? page.get("odata.nextLink").asText()
                                    : page.has("@odata.nextLink")
                                            ? page.get("@odata.nextLink").asText()
                                            : null;
        }
        return rows;
    }

    public Map<String, Object> getItem(String title, long id, String query) {
        return toMap(
                request(
                        HttpMethod.GET,
                        listPath(title)
                                + "/items("
                                + positive(id)
                                + ")"
                                + (query.isEmpty() ? "" : "?" + query),
                        null,
                        null,
                        false));
    }

    public Map<String, Object> createItem(String title, Map<String, Object> fields) {
        return toMap(
                request(
                        HttpMethod.POST,
                        listPath(title) + "/items",
                        typedFields(title, fields),
                        null,
                        true));
    }

    public void updateItem(String title, long id, Map<String, Object> fields, String eTag) {
        if (fields.isEmpty()) return;
        if (eTag == null || eTag.isBlank() || eTag.equals("*"))
            throw new IllegalArgumentException("An original ETag is required.");
        // MERGE changes only supplied fields. Use POST tunnelling supported by SharePoint REST.
        var headers = new HttpHeaders();
        headers.set("IF-MATCH", eTag);
        headers.set("X-HTTP-Method", "MERGE");
        exchange(
                HttpMethod.POST,
                listPath(title) + "/items(" + positive(id) + ")",
                typedFields(title, fields),
                headers);
    }

    public long appendInFolder(String title, String folderPath, String text) {
        var body =
                Map.of(
                        "listItemCreateInfo",
                        Map.of(
                                "FolderPath",
                                Map.of("DecodedUrl", absoluteFolder(folderPath)),
                                "UnderlyingObjectType",
                                0),
                        "formValues",
                        List.of(
                                Map.of("FieldName", "Title", "FieldValue", "Progress update"),
                                Map.of("FieldName", "EntryText", "FieldValue", text)),
                        "bNewDocumentUpdate",
                        false);
        JsonNode response =
                request(
                        HttpMethod.POST,
                        listPath(title) + "/AddValidateUpdateItemUsingPath",
                        body,
                        null,
                        true);
        JsonNode values =
                response.has("results")
                        ? response.get("results")
                        : response.has("AddValidateUpdateItemUsingPath")
                                ? response.path("AddValidateUpdateItemUsingPath").path("results")
                                : response.path("value");
        long id = 0;
        if (!values.isArray()) throw new AcceptedWriteException();
        for (JsonNode field : values) {
            if (field.path("HasException").asBoolean()) {
                throw new IllegalArgumentException(
                        field.path("ErrorMessage").asText("Progress entry validation failed."));
            }
            if (field.path("ItemId").asLong() > 0) id = field.path("ItemId").asLong();
            if (field.path("FieldName").asText().equalsIgnoreCase("Id"))
                id = field.path("FieldValue").asLong(id);
        }
        if (id <= 0) throw new AcceptedWriteException();
        return id;
    }

    public void ensureFolder(String folderPath) {
        try {
            JsonNode folder =
                    request(
                            HttpMethod.GET,
                            "web/GetFolderByServerRelativePath(decodedUrl='"
                                    + literal(folderPath)
                                    + "')?$select=Exists",
                            null,
                            null,
                            false);
            if (folder.path("Exists").asBoolean(true)) return;
        } catch (org.springframework.web.client.HttpClientErrorException e) {
            if (e.getStatusCode().value() != 404) throw e;
        }
        // Native caller permissions decide whether lazy creation is allowed; never elevate to
        // create it.
        request(
                HttpMethod.POST,
                "web/Folders/AddUsingPath(decodedUrl='" + literal(folderPath) + "')",
                Map.of(),
                null,
                false);
    }

    private Map<String, Object> typedFields(String title, Map<String, Object> fields) {
        var body = new LinkedHashMap<>(fields);
        body.put("__metadata", Map.of("type", listInfo(title).entityType()));
        return body;
    }

    private JsonNode request(
            HttpMethod method,
            String path,
            Object body,
            HttpHeaders headers,
            boolean acceptedWrite) {
        ResponseEntity<String> response = exchange(method, path, body, headers);
        try {
            JsonNode node = json.readTree(response.getBody());
            if (node == null || !node.isObject())
                throw new IllegalStateException("SharePoint returned an invalid response object.");
            node = node.has("d") ? node.get("d") : node;
            if (node == null || !node.isObject())
                throw new IllegalStateException("SharePoint returned an invalid response object.");
            String eTag = response.getHeaders().getETag();
            if (eTag != null && node instanceof ObjectNode object) object.put("odata.etag", eTag);
            return node;
        } catch (Exception e) {
            if (acceptedWrite) throw new AcceptedWriteException();
            throw new IllegalStateException("SharePoint returned an unreadable response.", e);
        }
    }

    private ResponseEntity<String> exchange(
            HttpMethod method, String path, Object body, HttpHeaders extra) {
        HttpHeaders headers = new HttpHeaders();
        headers.setBearerAuth(tokens.accessToken());
        headers.set("Accept", "application/json;odata=verbose");
        headers.set("Content-Type", "application/json;odata=verbose");
        if (extra != null) headers.addAll(extra);
        return rest.exchange(uri(path), method, new HttpEntity<>(body, headers), String.class);
    }

    private URI uri(String path) {
        String base = properties.getSiteUrl().replaceAll("/$", "") + "/_api/";
        URI candidate;
        if (path.startsWith("https://")) candidate = URI.create(path);
        else
            candidate =
                    URI.create(
                            base
                                    + UriUtils.encodePath(
                                            path.split("\\?", 2)[0], StandardCharsets.UTF_8)
                                    + (path.contains("?")
                                            ? "?"
                                                    + UriUtils.encodeQuery(
                                                            path.substring(path.indexOf('?') + 1),
                                                            StandardCharsets.UTF_8)
                                            : ""));
        URI allowed = URI.create(base);
        if (!"https".equals(candidate.getScheme())
                || !allowed.getAuthority().equalsIgnoreCase(candidate.getAuthority())
                || !candidate.getPath().startsWith(allowed.getPath())
                || candidate.getRawUserInfo() != null
                || !candidate.normalize().getPath().equals(candidate.getPath())) {
            throw new IllegalArgumentException("SharePoint continuation left the configured site.");
        }
        return candidate;
    }

    private String absoluteFolder(String path) {
        URI site = URI.create(properties.getSiteUrl());
        String sitePath = site.getPath().replaceAll("/$", "");
        if (!path.startsWith(sitePath + "/") || path.contains("/../"))
            throw new IllegalArgumentException("Invalid journal folder.");
        return site.getScheme() + "://" + site.getAuthority() + path;
    }

    private Map<String, Object> toMap(JsonNode node) {
        return json.convertValue(node, new TypeReference<>() {});
    }

    private static long positive(long id) {
        if (id <= 0) throw new IllegalArgumentException("Invalid item ID.");
        return id;
    }

    public static String literal(String value) {
        return value.replace("'", "''");
    }

    private String listPath(String title) {
        return "web/lists/GetByTitle('" + literal(title) + "')";
    }
}

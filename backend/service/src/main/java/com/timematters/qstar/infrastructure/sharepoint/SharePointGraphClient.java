package com.timematters.qstar.infrastructure.sharepoint;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.node.ObjectNode;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.client.HttpStatusCodeException;
import org.springframework.web.client.RestTemplate;

/**
 * Low-level Microsoft Graph client for reading/writing SharePoint list items. This is the
 * server-side counterpart to what the SPFx web part's SharePointDataService.ts does client-side
 * via PnPjs — the difference is this authenticates as the backend's own app identity
 * (client-credentials, Sites.Selected) rather than the signed-in user, since Q-Star's "thin
 * gateway" backend is the one talking to SharePoint, not the browser.
 *
 * <p>Deliberately built on plain RestTemplate calls against the Graph REST API, following the
 * exact same pattern already used by MicrosoftGraphClient (user lookups) elsewhere in this
 * package, rather than pulling in the full Microsoft Graph Java SDK — one fewer dependency whose
 * exact API surface this codebase would otherwise need to get right on the first try.
 */
@Service
public class SharePointGraphClient {

    private static final String GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
    private static final Logger log = LoggerFactory.getLogger(SharePointGraphClient.class);

    @Value("${qstar.sharepoint.tenant-id}")
    private String tenantId;

    @Value("${qstar.sharepoint.client-id}")
    private String clientId;

    @Value("${qstar.sharepoint.client-secret}")
    private String clientSecret;

    @Value("${qstar.sharepoint.site-hostname}")
    private String siteHostname;

    @Value("${qstar.sharepoint.site-path}")
    private String sitePath;

    private final RestTemplate rest = new RestTemplate();
    private final ObjectMapper objectMapper = new ObjectMapper();

    private String token = "";
    private String cachedSiteId;

    /** Resolves and caches the SharePoint site id for {@code siteHostname}/{@code sitePath}. */
    public String getSiteId() {
        if (cachedSiteId != null) {
            return cachedSiteId;
        }
        String url = GRAPH_BASE_URL + "/sites/" + siteHostname + ":" + sitePath;
        JsonNode response = exchange(HttpMethod.GET, url, null);
        cachedSiteId = response.get("id").asText();
        return cachedSiteId;
    }

    /** Resolves a list's id from its display name. */
    public String getListId(String listDisplayName) {
        String url =
                GRAPH_BASE_URL
                        + "/sites/"
                        + getSiteId()
                        + "/lists?$filter="
                        + urlEncode("displayName eq '" + listDisplayName + "'");
        JsonNode response = exchange(HttpMethod.GET, url, null);
        JsonNode values = response.get("value");
        if (values == null || values.isEmpty()) {
            throw new IllegalStateException(
                    "SharePoint list \"" + listDisplayName + "\" was not found on the configured site.");
        }
        return values.get(0).get("id").asText();
    }

    /** Returns every item's "fields" object for the given list, id included as "Id". */
    public List<Map<String, Object>> getAllItems(String listId) {
        List<Map<String, Object>> items = new ArrayList<>();
        String url = GRAPH_BASE_URL + "/sites/" + getSiteId() + "/lists/" + listId + "/items?expand=fields&$top=200";
        while (url != null) {
            JsonNode response = exchange(HttpMethod.GET, url, null);
            for (JsonNode item : response.get("value")) {
                items.add(flatten(item));
            }
            JsonNode nextLink = response.get("@odata.nextLink");
            url = nextLink != null ? nextLink.asText() : null;
        }
        return items;
    }

    /** Creates an item and returns its "fields" object (id included as "Id"). */
    public Map<String, Object> createItem(String listId, Map<String, Object> fields) {
        String url = GRAPH_BASE_URL + "/sites/" + getSiteId() + "/lists/" + listId + "/items";
        ObjectNode body = objectMapper.createObjectNode();
        body.set("fields", objectMapper.valueToTree(fields));
        JsonNode response = exchange(HttpMethod.POST, url, body);
        return flatten(response);
    }

    /** Updates only the given fields on an existing item. */
    public void updateItemFields(String listId, String itemId, Map<String, Object> fields) {
        if (fields.isEmpty()) {
            return;
        }
        String url =
                GRAPH_BASE_URL + "/sites/" + getSiteId() + "/lists/" + listId + "/items/" + itemId + "/fields";
        exchange(HttpMethod.PATCH, url, objectMapper.valueToTree(fields));
    }

    public void deleteItem(String listId, String itemId) {
        String url = GRAPH_BASE_URL + "/sites/" + getSiteId() + "/lists/" + listId + "/items/" + itemId;
        exchange(HttpMethod.DELETE, url, null);
    }

    /** Flattens a Graph list-item envelope ({id, fields:{...}}) into one fields map with "Id" added. */
    private Map<String, Object> flatten(JsonNode item) {
        JsonNode fields = item.has("fields") ? item.get("fields") : item;
        Map<String, Object> map = objectMapper.convertValue(fields, Map.class);
        if (item.has("id")) {
            map.put("Id", item.get("id").asText());
        }
        return map;
    }

    private String urlEncode(String value) {
        return URLEncoder.encode(value, StandardCharsets.UTF_8);
    }

    private JsonNode exchange(HttpMethod method, String url, Object body) {
        if (token.isEmpty()) {
            token = fetchOauthToken();
        }
        try {
            return doExchange(method, url, body, token);
        } catch (HttpStatusCodeException e) {
            if (e.getStatusCode() == HttpStatus.UNAUTHORIZED) {
                log.warn("Graph token expired/rejected, refreshing and retrying once.");
                token = fetchOauthToken();
                return doExchange(method, url, body, token);
            }
            throw e;
        }
    }

    private JsonNode doExchange(HttpMethod method, String url, Object body, String bearerToken) {
        HttpHeaders headers = new HttpHeaders();
        headers.setBearerAuth(bearerToken);
        headers.setContentType(MediaType.APPLICATION_JSON);
        HttpEntity<Object> entity = new HttpEntity<>(body, headers);
        return rest.exchange(url, method, entity, JsonNode.class).getBody();
    }

    private String fetchOauthToken() {
        String requestUrl = "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token";

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_FORM_URLENCODED);

        String parameterString =
                "client_id="
                        + urlEncode(clientId)
                        + "&scope="
                        + urlEncode("https://graph.microsoft.com/.default")
                        + "&client_secret="
                        + urlEncode(clientSecret)
                        + "&grant_type=client_credentials";

        HttpEntity<String> entity = new HttpEntity<>(parameterString, headers);
        JsonNode response = rest.exchange(requestUrl, HttpMethod.POST, entity, JsonNode.class).getBody();
        log.info("Graph app-only OAuth token fetched.");
        return response.get("access_token").asText();
    }
}

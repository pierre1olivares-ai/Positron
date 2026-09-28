package com.timematters.qstar.api.authentication;

import com.timematters.qstar.api.authentication.graph.GraphAuthDTO;
import com.timematters.qstar.api.authentication.graph.GraphUserInfoDTO;
import java.net.URLEncoder;
import java.util.Optional;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.*;
import org.springframework.stereotype.Service;
import org.springframework.web.client.HttpStatusCodeException;
import org.springframework.web.client.RestClientException;
import org.springframework.web.client.RestTemplate;

@Service
public class MicrosoftGraphClient {

    private String BASE_URL = "https://graph.microsoft.com/";
    private String AUTH_HEADER = "Authorization";
    private String AUTH_PREFIX = "Bearer ";

    @Value("${azure.activedirectory.tenant-id}")
    private String tenantId;

    @Value("${azure.activedirectory.client-id}")
    private String clientId;

    @Value("${azure.activedirectory.client-secret}")
    private String clientSecret;

    private String token = "";
    private static final Logger log = LoggerFactory.getLogger(MicrosoftGraphClient.class);

    public Optional<GraphUserInfoDTO> getUserInfo(String userId) throws Exception {
        ResponseEntity<GraphUserInfoDTO> response = getGraphUserInfoDTOResponseEntity(userId);
        if (response.getStatusCode() == HttpStatus.OK) {
            log.info("Fetched user infos for userID: " + userId);
            return Optional.of(response.getBody());
        } else {
            log.warn(
                    "Tried to fetch user Info for "
                            + userId
                            + "and got the following response:"
                            + response.toString());
            return Optional.empty();
        }
    }

    private ResponseEntity<GraphUserInfoDTO> getGraphUserInfoDTOResponseEntity(String userId)
            throws Exception {
        String userInfoUrl = BASE_URL + "v1.0/users/";
        ResponseEntity response;
        RestTemplate rest = new RestTemplate();
        HttpHeaders headers = new HttpHeaders();
        if (token.isEmpty()) {
            token = fetchOauthToken();
        }
        headers.add(AUTH_HEADER, AUTH_PREFIX + token);
        HttpEntity<?> entity = new HttpEntity<>(headers);

        try {
            response =
                    rest.exchange(
                            userInfoUrl + userId, HttpMethod.GET, entity, GraphUserInfoDTO.class);
        } catch (HttpStatusCodeException e) {
            if (e.getStatusCode() == HttpStatus.UNAUTHORIZED) {
                log.warn("Got Authorization Error:  " + e.toString());
                token = fetchOauthToken();
                rest = new RestTemplate();
                headers = new HttpHeaders();
                headers.add(AUTH_HEADER, AUTH_PREFIX + token);
                entity = new HttpEntity<>(headers);
                try {
                    response =
                            rest.exchange(
                                    userInfoUrl + userId,
                                    HttpMethod.GET,
                                    entity,
                                    GraphUserInfoDTO.class);
                } catch (RestClientException ex) {
                    throw e;
                }
            } else {
                throw e;
            }
        }
        return response;
    }

    private String fetchOauthToken() throws Exception {

        String rawScope = BASE_URL + ".default";
        String grant_type = "client_credentials";

        String encodedScope = URLEncoder.encode(rawScope, "UTF-8");
        String encodedSecret = URLEncoder.encode(clientSecret, "UTF-8");

        String requestUrl = "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token";

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_FORM_URLENCODED);

        String parameterString =
                "client_id="
                        + clientId
                        + "&scope="
                        + encodedScope
                        + "&client_secret="
                        + encodedSecret
                        + "&grant_type="
                        + grant_type;

        HttpEntity<?> entity = new HttpEntity<>(parameterString, headers);

        RestTemplate rest = new RestTemplate();
        HttpEntity<GraphAuthDTO> response =
                rest.exchange(requestUrl, HttpMethod.POST, entity, GraphAuthDTO.class);
        log.info("OAuthToken has been fetched.");

        return response.getBody().getAccess_token();
    }
}

package com.timematters.qstar.configuration.security;

import static org.junit.jupiter.api.Assertions.*;
import static org.springframework.test.web.client.match.MockRestRequestMatchers.*;
import static org.springframework.test.web.client.response.MockRestResponseCreators.*;

import com.azure.spring.cloud.autoconfigure.implementation.aad.configuration.properties.AadAuthenticationProperties;
import com.azure.spring.cloud.autoconfigure.implementation.aad.security.AadJwtBearerGrantRequestEntityConverter;
import com.timematters.qstar.infrastructure.sharepoint.SharePointProperties;
import java.util.List;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.support.DefaultListableBeanFactory;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.http.converter.FormHttpMessageConverter;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.oauth2.client.JwtBearerOAuth2AuthorizedClientProvider;
import org.springframework.security.oauth2.client.endpoint.DefaultJwtBearerTokenResponseClient;
import org.springframework.security.oauth2.client.http.OAuth2ErrorResponseErrorHandler;
import org.springframework.security.oauth2.client.registration.ClientRegistrationRepository;
import org.springframework.security.oauth2.core.http.converter.OAuth2AccessTokenResponseHttpMessageConverter;
import org.springframework.security.oauth2.server.resource.authentication.JwtAuthenticationToken;
import org.springframework.test.web.client.MockRestServiceServer;
import org.springframework.util.LinkedMultiValueMap;
import org.springframework.web.client.RestTemplate;

class OboTokenProviderTest {
    @AfterEach
    void clearCaller() {
        SecurityContextHolder.clearContext();
    }

    @Test
    void exchangeContainsCallerAssertionAndCachesOnlyWithinOneRequest() {
        var security = QstarSecurityTest.properties();
        var sharePoint = new SharePointProperties();
        sharePoint.setSiteUrl("https://example.sharepoint.com/sites/Quality");
        var azure = new AadAuthenticationProperties();
        azure.getCredential().setClientSecret("test-secret");
        var registrations =
                new WebSecurityConfig().clientRegistrationRepository(azure, security, sharePoint);
        var rest =
                new RestTemplate(
                        List.of(
                                new FormHttpMessageConverter(),
                                new OAuth2AccessTokenResponseHttpMessageConverter()));
        rest.setErrorHandler(new OAuth2ErrorResponseErrorHandler());
        var server = MockRestServiceServer.bindTo(rest).build();
        var tokenClient = new DefaultJwtBearerTokenResponseClient();
        tokenClient.setRestOperations(rest);
        tokenClient.setRequestEntityConverter(new AadJwtBearerGrantRequestEntityConverter());
        var oauth = new JwtBearerOAuth2AuthorizedClientProvider();
        oauth.setAccessTokenResponseClient(tokenClient);
        var beans = new DefaultListableBeanFactory();
        beans.registerSingleton("oauth", oauth);
        beans.registerSingleton("registrations", registrations);
        var backend = new BackendProperties();
        backend.setEnabled(true);
        for (String caller : List.of("caller-one", "caller-two")) {
            var expected = new LinkedMultiValueMap<String, String>();
            expected.add("grant_type", "urn:ietf:params:oauth:grant-type:jwt-bearer");
            expected.add("client_id", QstarSecurityTest.CLIENT);
            expected.add("client_secret", "test-secret");
            expected.add("assertion", caller);
            expected.add("scope", "https://example.sharepoint.com/.default");
            expected.add("requested_token_use", "on_behalf_of");
            server.expect(
                            requestTo(
                                    "https://login.microsoftonline.com/"
                                            + QstarSecurityTest.TENANT
                                            + "/oauth2/v2.0/token"))
                    .andExpect(method(HttpMethod.POST))
                    .andExpect(content().formData(expected))
                    .andRespond(
                            withSuccess(
                                    "{\"access_token\":\"sp-"
                                            + caller
                                            + "\",\"token_type\":\"Bearer\",\"expires_in\":3600}",
                                    MediaType.APPLICATION_JSON));
        }
        for (String caller : List.of("caller-one", "caller-two")) {
            SecurityContextHolder.getContext()
                    .setAuthentication(
                            new JwtAuthenticationToken(
                                    QstarSecurityTest.token(caller).build(), List.of()));
            var requestProvider =
                    new OboSharePointAccessTokenProvider(
                            backend,
                            beans.getBeanProvider(JwtBearerOAuth2AuthorizedClientProvider.class),
                            beans.getBeanProvider(ClientRegistrationRepository.class));
            assertEquals("sp-" + caller, requestProvider.accessToken());
            assertEquals("sp-" + caller, requestProvider.accessToken());
        }
        server.verify();
    }
}

package com.timematters.qstar.configuration.security;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.*;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

import com.nimbusds.jose.JWSAlgorithm;
import com.nimbusds.jose.JWSHeader;
import com.nimbusds.jose.crypto.RSASSASigner;
import com.nimbusds.jwt.JWTClaimsSet;
import com.nimbusds.jwt.SignedJWT;
import com.timematters.qstar.infrastructure.sharepoint.SharePointRestClient;
import com.timematters.qstar.infrastructure.sharepoint.SharePointUser;
import java.security.KeyPair;
import java.security.KeyPairGenerator;
import java.security.interfaces.RSAPublicKey;
import java.time.Instant;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.test.context.TestConfiguration;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Import;
import org.springframework.context.annotation.Primary;
import org.springframework.security.oauth2.client.registration.ClientRegistrationRepository;
import org.springframework.security.oauth2.core.AuthorizationGrantType;
import org.springframework.security.oauth2.jwt.JwtDecoder;
import org.springframework.security.oauth2.jwt.NimbusJwtDecoder;
import org.springframework.test.web.servlet.MockMvc;

@SpringBootTest(
        properties = {
            "qstar.backend.enabled=true",
            "qstar.security.tenant-id=11111111-1111-1111-1111-111111111111",
            "qstar.security.client-id=22222222-2222-2222-2222-222222222222",
            "qstar.security.allowed-origins=https://example.sharepoint.com",
            "qstar.sharepoint.site-url=https://example.sharepoint.com/sites/Quality",
            "spring.cloud.azure.active-directory.credential.client-secret=local-test-only"
        })
@AutoConfigureMockMvc
@Import(EnabledBackendSecurityTest.LocalSigningKeys.class)
class EnabledBackendSecurityTest {
    static final KeyPair SIGNING_KEY = keyPair();
    @Autowired MockMvc mvc;
    @Autowired ClientRegistrationRepository registrations;
    @MockBean SharePointRestClient sharePoint;

    @TestConfiguration(proxyBeanMethods = false)
    static class LocalSigningKeys {
        @Bean
        @Primary
        JwtDecoder localDecoder(QstarSecurityProperties properties) {
            return WebSecurityConfig.configureValidator(
                    NimbusJwtDecoder.withPublicKey((RSAPublicKey) SIGNING_KEY.getPublic()).build(),
                    properties);
        }
    }

    static KeyPair keyPair() {
        try {
            var generator = KeyPairGenerator.getInstance("RSA");
            generator.initialize(2048);
            return generator.generateKeyPair();
        } catch (Exception exception) {
            throw new IllegalStateException(exception);
        }
    }

    static String bearer(String role) throws Exception {
        return bearer(role, Map.of(), SIGNING_KEY);
    }

    static String bearer(String role, Map<String, Object> replacements, KeyPair key)
            throws Exception {
        var claims =
                new JWTClaimsSet.Builder()
                        .issuer(
                                "https://login.microsoftonline.com/"
                                        + QstarSecurityTest.TENANT
                                        + "/v2.0")
                        .audience(QstarSecurityTest.CLIENT)
                        .subject(QstarSecurityTest.USER)
                        .issueTime(Date.from(Instant.now().minusSeconds(10)))
                        .expirationTime(Date.from(Instant.now().plusSeconds(600)))
                        .claim("tid", QstarSecurityTest.TENANT)
                        .claim("oid", QstarSecurityTest.USER)
                        .claim("scp", "user_impersonation")
                        .claim("roles", List.of(role));
        replacements.forEach(claims::claim);
        var token = new SignedJWT(new JWSHeader(JWSAlgorithm.RS256), claims.build());
        token.sign(new RSASSASigner(key.getPrivate()));
        return "Bearer " + token.serialize();
    }

    @BeforeEach
    void identity() {
        when(sharePoint.currentUser())
                .thenReturn(new SharePointUser(42, "Actual Caller", "actual@example.com"));
    }

    @Test
    void requiresBearerAndReturnsAuthoritativeIdentityAndConfiguredConnection() throws Exception {
        mvc.perform(get("/api/v1/me")).andExpect(status().isUnauthorized());
        mvc.perform(get("/api/v1/me").header("Authorization", bearer("QStar.Owner")))
                .andExpect(status().isOk())
                .andExpect(jsonPath("$.role").value("owner"))
                .andExpect(jsonPath("$.user.userId").value(42))
                .andExpect(jsonPath("$.user.email").value("actual@example.com"))
                .andExpect(
                        jsonPath("$.connection.siteUrl")
                                .value("https://example.sharepoint.com/sites/Quality"));
        assertEquals(
                AuthorizationGrantType.JWT_BEARER,
                registrations.findByRegistrationId("sharepoint").getAuthorizationGrantType());
        assertNull(registrations.findByRegistrationId("graph"));
    }

    @Test
    void actualBearerFilterRejectsExpiredMissingExpiryWrongSignatureAndApplicationTokens()
            throws Exception {
        var missingExpiry = new HashMap<String, Object>();
        missingExpiry.put("exp", null);
        for (String token :
                List.of(
                        bearer(
                                "QStar.Admin",
                                Map.of("exp", Date.from(Instant.now().minusSeconds(600))),
                                SIGNING_KEY),
                        bearer("QStar.Admin", missingExpiry, SIGNING_KEY),
                        bearer("QStar.Admin", Map.of(), keyPair()),
                        bearer("QStar.Admin", Map.of("idtyp", "app"), SIGNING_KEY),
                        bearer("QStar.Admin", Map.of("tid", QstarSecurityTest.CLIENT), SIGNING_KEY),
                        bearer(
                                "QStar.Admin",
                                Map.of("aud", "https://graph.microsoft.com"),
                                SIGNING_KEY))) {
            mvc.perform(get("/api/v1/me").header("Authorization", token))
                    .andExpect(status().isUnauthorized());
        }
        verify(sharePoint, never()).currentUser();
    }

    @Test
    void readerCannotRunAdminOperationsAndCorsUsesExactTenantOrigin() throws Exception {
        String reader = bearer("Unmapped.Admin");
        mvc.perform(get("/api/v1/diagnostics").header("Authorization", reader))
                .andExpect(status().isForbidden());
        mvc.perform(
                        put("/api/v1/settings")
                                .header("Authorization", reader)
                                .contentType("application/json")
                                .content("{}"))
                .andExpect(status().isForbidden());
        verify(sharePoint, never()).getSiteId();
        mvc.perform(
                        options("/api/v1/issues")
                                .header("Origin", "https://example.sharepoint.com")
                                .header("Access-Control-Request-Method", "PATCH")
                                .header(
                                        "Access-Control-Request-Headers",
                                        "Authorization,Content-Type,If-Match"))
                .andExpect(status().isOk())
                .andExpect(
                        header().string(
                                        "Access-Control-Allow-Origin",
                                        "https://example.sharepoint.com"));
        mvc.perform(
                        options("/api/v1/issues")
                                .header("Origin", "https://example.sharepoint.com.evil.test")
                                .header("Access-Control-Request-Method", "PATCH"))
                .andExpect(status().isForbidden());
        mvc.perform(
                        get("/api/v1/me")
                                .header("Origin", "https://example.sharepoint.com")
                                .header("Authorization", reader))
                .andExpect(jsonPath("$.role").value("reader"))
                .andExpect(
                        header().string(
                                        "Access-Control-Expose-Headers",
                                        "ETag, Location, X-QStar-Reference, X-QStar-Entry-Id"));
    }
}

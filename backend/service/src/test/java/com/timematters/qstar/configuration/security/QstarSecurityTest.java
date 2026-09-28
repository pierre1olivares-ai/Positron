package com.timematters.qstar.configuration.security;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

import com.timematters.qstar.api.controller.DiagnosticsApiController;
import com.timematters.qstar.api.controller.SettingsApiController;
import com.timematters.qstar.infrastructure.sharepoint.SettingsRepository;
import com.timematters.qstar.infrastructure.sharepoint.SharePointProperties;
import com.timematters.qstar.service.DiagnosticsService;
import java.time.Instant;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.security.oauth2.core.DelegatingOAuth2TokenValidator;
import org.springframework.security.oauth2.jwt.Jwt;
import org.springframework.security.oauth2.jwt.JwtValidators;

class QstarSecurityTest {
    static final String TENANT = "11111111-1111-1111-1111-111111111111";
    static final String CLIENT = "22222222-2222-2222-2222-222222222222";
    static final String USER = "33333333-3333-3333-3333-333333333333";

    static QstarSecurityProperties properties() {
        var properties = new QstarSecurityProperties();
        properties.setTenantId(TENANT);
        properties.setClientId(CLIENT);
        properties.setAppIdUri("api://" + CLIENT);
        properties.setAllowedOrigins(List.of("https://example.sharepoint.com"));
        return properties;
    }

    static Jwt.Builder token(String value) {
        return Jwt.withTokenValue(value)
                .header("alg", "RS256")
                .issuer("https://login.microsoftonline.com/" + TENANT + "/v2.0")
                .audience(List.of(CLIENT))
                .subject(USER)
                .issuedAt(Instant.now().minusSeconds(10))
                .expiresAt(Instant.now().plusSeconds(600))
                .claim("tid", TENANT)
                .claim("oid", USER)
                .claim("scp", "user_impersonation");
    }

    @Test
    void rejectsOtherTenantsWrongAudienceApplicationAndMissingScopeTokens() {
        var validator = new QstarJwtValidator(properties());
        assertFalse(validator.validate(token("valid").build()).hasErrors());
        var invalid =
                List.of(
                        token("wrong-tenant").claim("tid", CLIENT).build(),
                        token("wrong-issuer")
                                .issuer("https://login.microsoftonline.com/" + CLIENT + "/v2.0")
                                .build(),
                        token("wrong-aud").audience(List.of("https://graph.microsoft.com")).build(),
                        token("multiple-aud").audience(List.of(CLIENT, "another-api")).build(),
                        token("app")
                                .claim("idtyp", "app")
                                .claim("roles", List.of("QStar.Admin"))
                                .build(),
                        token("no-scope").claims(claims -> claims.remove("scp")).build(),
                        token("no-expiry").claims(claims -> claims.remove("exp")).build(),
                        token("scope-substring").claim("scp", "other_user_impersonation").build(),
                        token("no-user").claims(claims -> claims.remove("oid")).build());
        invalid.forEach(
                jwt -> assertTrue(validator.validate(jwt).hasErrors(), jwt.getTokenValue()));
        var combined =
                new DelegatingOAuth2TokenValidator<>(JwtValidators.createDefault(), validator);
        assertTrue(
                combined.validate(
                                token("expired")
                                        .issuedAt(Instant.now().minusSeconds(4000))
                                        .expiresAt(Instant.now().minusSeconds(1000))
                                        .build())
                        .hasErrors());
    }

    @Test
    void rolesUseOnlyExplicitAppRoleValuesAndConfigurationCannotAliasThem() {
        var properties = properties();
        var roles = new AppRoleResolver(properties);
        assertEquals(
                "reader",
                roles.role(
                        token("unknown")
                                .claim("roles", List.of("Admin", "Global Administrator"))
                                .claim("groups", List.of("QStar.Admin"))
                                .build()));
        assertEquals(
                "owner", roles.role(token("owner").claim("roles", List.of("QStar.Owner")).build()));
        assertEquals(
                "qm",
                roles.role(token("qm").claim("roles", List.of("QStar.QM", "QStar.Owner")).build()));
        assertEquals(
                "admin", roles.role(token("admin").claim("roles", List.of("QStar.Admin")).build()));
        properties.setOwnerRole(properties.getAdminRole());
        assertThrows(IllegalArgumentException.class, properties::validate);
    }

    @Test
    void ownerAndReaderCannotReachAdminMutationsOrDiagnostics() throws Exception {
        var settings = mock(SettingsRepository.class);
        var diagnostics = mock(DiagnosticsService.class);
        for (String role : List.of("reader", "owner", "qm")) {
            var policy =
                    new AuthorizationPolicy(
                            () -> new CallerContext(role, 7, "Caller", "caller@example.com"));
            var settingsController = new SettingsApiController(null, settings, policy);
            var diagnosticsController = new DiagnosticsApiController(null, diagnostics, policy);
            assertThrows(AccessDeniedException.class, () -> settingsController.saveSettings(null));
            assertThrows(AccessDeniedException.class, diagnosticsController::getDiagnostics);
        }
        verifyNoInteractions(settings, diagnostics);
        var policy =
                new AuthorizationPolicy(
                        () -> new CallerContext("owner", 7, "Caller", "caller@example.com"));
        assertThrows(AccessDeniedException.class, policy::requireManager);
        assertEquals(7, policy.requireContributor().sharePointUserId());
    }

    @Test
    void siteAndCorsConfigurationRejectsUnpinnedOrWildcardTargets() {
        var site = new SharePointProperties();
        site.setSiteUrl("https://example.sharepoint.com/sites/Quality/");
        site.validate();
        assertEquals("https://example.sharepoint.com/.default", site.tokenScope());
        for (String invalid :
                List.of(
                        "http://example.sharepoint.com/sites/Quality",
                        "https://example.sharepoint.com.evil.test/",
                        "https://user@example.sharepoint.com/sites/Quality",
                        "https://example.sharepoint.com/sites/Quality?next=other")) {
            site.setSiteUrl(invalid);
            assertThrows(IllegalArgumentException.class, site::validate);
        }
        var properties = properties();
        properties.setAllowedOrigins(List.of("https://*.sharepoint.com"));
        assertThrows(IllegalArgumentException.class, properties::validate);
        properties.setAllowedOrigins(List.of("https://example.sharepoint.com/"));
        assertThrows(IllegalArgumentException.class, properties::validate);
    }
}

package com.timematters.qstar.configuration.security;

import org.springframework.beans.factory.ObjectProvider;
import org.springframework.http.HttpStatus;
import org.springframework.security.oauth2.client.JwtBearerOAuth2AuthorizedClientProvider;
import org.springframework.security.oauth2.client.OAuth2AuthorizationContext;
import org.springframework.security.oauth2.client.registration.ClientRegistrationRepository;
import org.springframework.security.oauth2.core.OAuth2AuthorizationException;
import org.springframework.stereotype.Component;
import org.springframework.web.context.annotation.RequestScope;
import org.springframework.web.server.ResponseStatusException;

/** A downstream token is never shared between HTTP requests or callers. */
@Component
@RequestScope
public class OboSharePointAccessTokenProvider implements SharePointAccessTokenProvider {
    private final BackendProperties backend;
    private final ObjectProvider<JwtBearerOAuth2AuthorizedClientProvider> providers;
    private final ObjectProvider<ClientRegistrationRepository> registrations;
    private String token;

    public OboSharePointAccessTokenProvider(
            BackendProperties backend,
            ObjectProvider<JwtBearerOAuth2AuthorizedClientProvider> providers,
            ObjectProvider<ClientRegistrationRepository> registrations) {
        this.backend = backend;
        this.providers = providers;
        this.registrations = registrations;
    }

    @Override
    public String accessToken() {
        if (!backend.isEnabled())
            throw new ResponseStatusException(
                    HttpStatus.SERVICE_UNAVAILABLE, "Q-Star backend is disabled");
        var principal = AppRoleResolver.authentication();
        if (token == null) {
            try {
                var registration = registrations.getObject().findByRegistrationId("sharepoint");
                var context =
                        OAuth2AuthorizationContext.withClientRegistration(registration)
                                .principal(principal)
                                .build();
                var authorized = providers.getObject().authorize(context);
                if (authorized == null)
                    throw new ResponseStatusException(
                            HttpStatus.UNAUTHORIZED,
                            "Delegated SharePoint authorization is required");
                token = authorized.getAccessToken().getTokenValue();
            } catch (OAuth2AuthorizationException exception) {
                // Never return/log assertions, secrets or raw token-endpoint response details.
                throw new ResponseStatusException(
                        HttpStatus.UNAUTHORIZED,
                        "Delegated SharePoint authorization failed; check consent and sign in again");
            }
        }
        return token;
    }
}

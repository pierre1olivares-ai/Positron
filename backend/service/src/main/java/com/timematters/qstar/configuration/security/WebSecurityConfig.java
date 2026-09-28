package com.timematters.qstar.configuration.security;

import com.azure.spring.cloud.autoconfigure.implementation.aad.configuration.properties.AadAuthenticationProperties;
import com.timematters.qstar.infrastructure.sharepoint.SharePointProperties;
import java.util.List;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.http.HttpMethod;
import org.springframework.security.config.Customizer;
import org.springframework.security.config.annotation.method.configuration.EnableMethodSecurity;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.http.SessionCreationPolicy;
import org.springframework.security.oauth2.client.registration.ClientRegistration;
import org.springframework.security.oauth2.client.registration.ClientRegistrationRepository;
import org.springframework.security.oauth2.client.registration.InMemoryClientRegistrationRepository;
import org.springframework.security.oauth2.core.AuthorizationGrantType;
import org.springframework.security.oauth2.core.ClientAuthenticationMethod;
import org.springframework.security.oauth2.core.DelegatingOAuth2TokenValidator;
import org.springframework.security.oauth2.jwt.JwtDecoder;
import org.springframework.security.oauth2.jwt.JwtValidators;
import org.springframework.security.oauth2.jwt.NimbusJwtDecoder;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.security.web.authentication.AnonymousAuthenticationFilter;
import org.springframework.util.Assert;
import org.springframework.web.cors.CorsConfiguration;
import org.springframework.web.cors.CorsConfigurationSource;
import org.springframework.web.cors.UrlBasedCorsConfigurationSource;

@Configuration(proxyBeanMethods = false)
@EnableMethodSecurity
@EnableConfigurationProperties({
    BackendProperties.class,
    QstarSecurityProperties.class,
    SharePointProperties.class
})
public class WebSecurityConfig {
    @Bean
    public SecurityFilterChain filterChain(
            HttpSecurity http,
            BackendProperties backend,
            QstarSecurityProperties security,
            SharePointProperties sharePoint)
            throws Exception {
        http.cors(Customizer.withDefaults())
                .csrf(csrf -> csrf.disable())
                .sessionManagement(
                        session -> session.sessionCreationPolicy(SessionCreationPolicy.STATELESS))
                .requestCache(cache -> cache.disable());
        if (backend.isEnabled()) {
            security.validate();
            sharePoint.validate();
            http.oauth2ResourceServer(oauth -> oauth.jwt(Customizer.withDefaults()));
            http.authorizeHttpRequests(
                    authorize ->
                            authorize
                                    .requestMatchers(HttpMethod.OPTIONS, "/**")
                                    .permitAll()
                                    .requestMatchers("/api/**")
                                    .authenticated()
                                    .anyRequest()
                                    .denyAll());
        } else {
            // The disabled service starts without credentials and cannot expose functional APIs.
            http.addFilterBefore(
                    (request, response, chain) -> {
                        var httpResponse = (jakarta.servlet.http.HttpServletResponse) response;
                        httpResponse.setStatus(503);
                        httpResponse.setContentType("application/json");
                        httpResponse
                                .getWriter()
                                .write(
                                        "{\"code\":\"BACKEND_DISABLED\",\"message\":\"The Q-Star backend is not enabled\"}");
                    },
                    AnonymousAuthenticationFilter.class);
            http.authorizeHttpRequests(authorize -> authorize.anyRequest().denyAll());
        }
        return http.build();
    }

    @Bean
    @ConditionalOnProperty(name = "qstar.backend.enabled", havingValue = "true")
    public JwtDecoder jwtDecoder(QstarSecurityProperties properties) {
        properties.validate();
        NimbusJwtDecoder decoder =
                NimbusJwtDecoder.withJwkSetUri(
                                "https://login.microsoftonline.com/"
                                        + properties.getTenantId()
                                        + "/discovery/v2.0/keys")
                        .build();
        return configureValidator(decoder, properties);
    }

    static JwtDecoder configureValidator(
            NimbusJwtDecoder decoder, QstarSecurityProperties properties) {
        decoder.setJwtValidator(
                new DelegatingOAuth2TokenValidator<>(
                        JwtValidators.createDefault(), new QstarJwtValidator(properties)));
        return decoder;
    }

    @Bean
    @ConditionalOnProperty(name = "qstar.backend.enabled", havingValue = "true")
    public ClientRegistrationRepository clientRegistrationRepository(
            AadAuthenticationProperties aad,
            QstarSecurityProperties security,
            SharePointProperties sharePoint) {
        security.validate();
        sharePoint.validate();
        var credential = aad.getCredential();
        boolean certificate =
                credential.getClientCertificatePath() != null
                        && !credential.getClientCertificatePath().isBlank();
        Assert.isTrue(
                certificate
                        || (credential.getClientSecret() != null
                                && !credential.getClientSecret().isBlank()),
                "An OBO certificate or client secret is required when the backend is enabled");
        if (certificate) {
            Assert.hasText(
                    credential.getClientCertificatePassword(),
                    "Azure Spring 5.17.1 requires a password-protected PFX/P12 for OBO certificate authentication");
        }
        // Only one downstream audience and one grant are registered. No application grant fallback.
        ClientRegistration registration =
                ClientRegistration.withRegistrationId("sharepoint")
                        .clientId(security.getClientId())
                        .clientSecret(credential.getClientSecret())
                        .clientAuthenticationMethod(
                                certificate
                                        ? ClientAuthenticationMethod.PRIVATE_KEY_JWT
                                        : ClientAuthenticationMethod.CLIENT_SECRET_POST)
                        .authorizationGrantType(AuthorizationGrantType.JWT_BEARER)
                        .tokenUri(
                                "https://login.microsoftonline.com/"
                                        + security.getTenantId()
                                        + "/oauth2/v2.0/token")
                        .scope(sharePoint.tokenScope())
                        .build();
        return new InMemoryClientRegistrationRepository(registration);
    }

    @Bean
    public CorsConfigurationSource corsConfigurationSource(QstarSecurityProperties properties) {
        CorsConfiguration cors = new CorsConfiguration();
        cors.setAllowedOrigins(properties.getAllowedOrigins());
        cors.setAllowedMethods(List.of("GET", "POST", "PATCH", "PUT", "OPTIONS"));
        cors.setAllowedHeaders(List.of("Authorization", "Content-Type", "Accept", "If-Match"));
        cors.setExposedHeaders(
                List.of("ETag", "Location", "X-QStar-Reference", "X-QStar-Entry-Id"));
        cors.setAllowCredentials(false);
        UrlBasedCorsConfigurationSource source = new UrlBasedCorsConfigurationSource();
        source.registerCorsConfiguration("/**", cors);
        return source;
    }
}

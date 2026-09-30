package com.timematters.qstar.configuration.security;

import java.util.Arrays;
import java.util.List;
import java.util.UUID;
import org.springframework.security.oauth2.core.OAuth2Error;
import org.springframework.security.oauth2.core.OAuth2TokenValidator;
import org.springframework.security.oauth2.core.OAuth2TokenValidatorResult;
import org.springframework.security.oauth2.jwt.Jwt;

/** Additional checks after cryptographic signature and timestamp validation. */
public final class QstarJwtValidator implements OAuth2TokenValidator<Jwt> {
    private final QstarSecurityProperties properties;

    public QstarJwtValidator(QstarSecurityProperties properties) {
        this.properties = properties;
    }

    @Override
    public OAuth2TokenValidatorResult validate(Jwt jwt) {
        try {
            String tenant = properties.getTenantId();
            String issuer = jwt.getClaimAsString("iss");
            boolean issuerValid =
                    ("https://login.microsoftonline.com/" + tenant + "/v2.0").equals(issuer)
                            || ("https://sts.windows.net/" + tenant + "/").equals(issuer);
            List<String> audiences = jwt.getAudience();
            boolean audienceValid =
                    audiences != null
                            && audiences.size() == 1
                            && (properties.getClientId().equals(audiences.get(0))
                                    || (!properties.getAppIdUri().isBlank()
                                            && properties.getAppIdUri().equals(audiences.get(0))));
            String scopes = jwt.getClaimAsString("scp");
            boolean scopeValid =
                    scopes != null
                            && Arrays.asList(scopes.split("\\s+"))
                                    .contains(properties.getRequiredScope());
            UUID.fromString(jwt.getClaimAsString("oid"));
            if (jwt.getExpiresAt() != null
                    && issuerValid
                    && tenant.equals(jwt.getClaimAsString("tid"))
                    && audienceValid
                    && scopeValid
                    && !"app".equals(jwt.getClaimAsString("idtyp"))) {
                return OAuth2TokenValidatorResult.success();
            }
        } catch (RuntimeException ignored) {
            // Malformed/missing claims fail closed; never reflect token contents in errors.
        }
        return OAuth2TokenValidatorResult.failure(
                new OAuth2Error(
                        "invalid_token",
                        "A delegated Q-Star token for the configured tenant, audience and scope is required",
                        null));
    }
}

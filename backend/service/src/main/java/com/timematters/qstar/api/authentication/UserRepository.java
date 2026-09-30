package com.timematters.qstar.api.authentication;

import com.timematters.qstar.api.authentication.user.AuthenticatedUser;
import com.timematters.qstar.api.authentication.user.UserId;
import java.util.Optional;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.oauth2.jwt.Jwt;

// NOTE: the template's version of this class was written against
// `com.microsoft.azure.spring.autoconfigure.aad.UserPrincipal`, from the
// legacy `azure-spring-boot-starter-active-directory` library. That library
// isn't on the classpath here — build.gradle instead wires up the newer
// `com.azure.spring:spring-cloud-azure-starter-active-directory`, which is
// also what configuration/security/WebSecurityConfig.java's
// AadResourceServerWebSecurityConfigurerAdapter comes from. Under that
// (OAuth2 Resource Server / JWT bearer) setup, the authenticated principal is
// a standard Spring Security `Jwt`, not a `UserPrincipal` — rewritten below
// to extract the same "oid" claim from that instead.
// Unwired legacy template example. CurrentUserProvider is the active delegated identity service.
public class UserRepository {

    private MicrosoftGraphClient graphClient;

    @Autowired
    public UserRepository(MicrosoftGraphClient microsoftGraphClient) {
        this.graphClient = microsoftGraphClient;
    }

    public Optional<AuthenticatedUser> getUserInfo(UserId userId) throws Exception {
        return graphClient.getUserInfo(userId.toString()).map(user -> user.toUser());
    }

    public UserId currentUserId() {
        Jwt jwt = getJwt();

        String userId = getUserId(jwt);

        return new UserId(userId);
    }

    public AuthenticatedUser getCurrentUser() throws Exception {
        return getUserInfo(currentUserId()).orElseThrow();
    }

    private String getUserId(Jwt jwt) {
        return jwt.getClaimAsString("oid");
    }

    private Jwt getJwt() {
        return (Jwt) SecurityContextHolder.getContext().getAuthentication().getPrincipal();
    }
}

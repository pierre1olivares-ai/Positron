package com.timematters.qstar.configuration.security;

import java.util.Set;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.stereotype.Component;

@Component
public class AuthorizationPolicy {
    private final CurrentUserProvider users;

    public AuthorizationPolicy(CurrentUserProvider users) {
        this.users = users;
    }

    public CallerContext requireAdmin() {
        return require(Set.of("admin"));
    }

    public CallerContext requireManager() {
        return require(Set.of("admin", "qm"));
    }

    public CallerContext requireContributor() {
        return require(Set.of("admin", "qm", "owner"));
    }

    private CallerContext require(Set<String> roles) {
        CallerContext caller = users.currentUser();
        if (!roles.contains(caller.role()))
            throw new AccessDeniedException("This operation is not permitted for your Q-Star role");
        return caller;
    }
}

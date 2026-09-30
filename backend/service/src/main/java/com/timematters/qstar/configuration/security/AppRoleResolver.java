package com.timematters.qstar.configuration.security;

import java.util.Collection;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.security.core.Authentication;
import org.springframework.security.core.context.SecurityContextHolder;
import org.springframework.security.oauth2.jwt.Jwt;
import org.springframework.stereotype.Component;

@Component
public class AppRoleResolver {
    private final QstarSecurityProperties properties;

    public AppRoleResolver(QstarSecurityProperties properties) {
        this.properties = properties;
    }

    public static Authentication authentication() {
        Authentication authentication = SecurityContextHolder.getContext().getAuthentication();
        if (authentication == null
                || !authentication.isAuthenticated()
                || !(authentication.getPrincipal() instanceof Jwt)) {
            throw new AccessDeniedException("A verified delegated caller is required");
        }
        return authentication;
    }

    public String role() {
        return role((Jwt) authentication().getPrincipal());
    }

    public String role(Jwt jwt) {
        Object claim = jwt.getClaims().get("roles");
        if (!(claim instanceof Collection<?> roles)) {
            return "reader";
        }
        if (roles.contains(properties.getAdminRole())) return "admin";
        if (roles.contains(properties.getQmRole())) return "qm";
        if (roles.contains(properties.getOwnerRole())) return "owner";
        return "reader";
    }
}

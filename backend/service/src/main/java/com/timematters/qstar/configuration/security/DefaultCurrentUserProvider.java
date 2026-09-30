package com.timematters.qstar.configuration.security;

import com.timematters.qstar.infrastructure.sharepoint.SharePointRestClient;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.stereotype.Component;
import org.springframework.web.context.annotation.RequestScope;

@Component
@RequestScope
public class DefaultCurrentUserProvider implements CurrentUserProvider {
    private final SharePointRestClient sharePoint;
    private final AppRoleResolver roles;
    private CallerContext caller;

    public DefaultCurrentUserProvider(SharePointRestClient sharePoint, AppRoleResolver roles) {
        this.sharePoint = sharePoint;
        this.roles = roles;
    }

    @Override
    public CallerContext currentUser() {
        if (caller == null) {
            String role = roles.role();
            var user = sharePoint.currentUser();
            if (user == null || user.id() <= 0)
                throw new AccessDeniedException("SharePoint caller identity is unavailable");
            caller = new CallerContext(role, user.id(), user.displayName(), user.email());
        }
        return caller;
    }
}

package com.timematters.qstar.api.controller;

import com.timematters.qstar.configuration.security.CurrentUserProvider;
import com.timematters.qstar.infrastructure.sharepoint.SharePointProperties;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("${openapi.qStarIssueManager.base-path:/api/v1}")
public class MeApiController {
    private final CurrentUserProvider users;
    private final SharePointProperties sharePoint;

    public MeApiController(CurrentUserProvider users, SharePointProperties sharePoint) {
        this.users = users;
        this.sharePoint = sharePoint;
    }

    @GetMapping("/me")
    public MeResponse me() {
        var caller = users.currentUser();
        return new MeResponse(
                caller.role(),
                "backend",
                new User(caller.sharePointUserId(), caller.displayName(), caller.email()),
                new Connection(
                        sharePoint.getSiteUrl(),
                        sharePoint.getIssuesListName(),
                        sharePoint.getProgressListName(),
                        sharePoint.getConfigListName(),
                        sharePoint.isBetaAccessMode()));
    }

    public record MeResponse(String role, String source, User user, Connection connection) {}

    public record User(long userId, String displayName, String email) {}

    public record Connection(
            String siteUrl,
            String issuesListName,
            String progressListName,
            String configListName,
            boolean betaAccessMode) {}
}

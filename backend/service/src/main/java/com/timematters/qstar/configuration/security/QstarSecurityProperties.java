package com.timematters.qstar.configuration.security;

import java.net.URI;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.UUID;
import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.util.Assert;

@Data
@ConfigurationProperties("qstar.security")
public class QstarSecurityProperties {
    private String tenantId = "";
    private String clientId = "";
    private String appIdUri = "";
    private String requiredScope = "user_impersonation";
    private List<String> allowedOrigins = new ArrayList<>();
    private String adminRole = "QStar.Admin";
    private String qmRole = "QStar.QM";
    private String ownerRole = "QStar.Owner";
    private String readerRole = "QStar.Reader";

    public void validate() {
        UUID.fromString(tenantId);
        UUID.fromString(clientId);
        Assert.hasText(requiredScope, "qstar.security.required-scope is required");
        Assert.isTrue(!requiredScope.contains(" "), "Configure one required delegated API scope");
        List<String> roles = List.of(adminRole, qmRole, ownerRole, readerRole);
        roles.forEach(role -> Assert.hasText(role, "Each Q-Star app role must be configured"));
        Assert.isTrue(
                new HashSet<>(roles).size() == 4, "Q-Star app-role mappings must be distinct");
        Assert.notEmpty(
                allowedOrigins, "Configure the exact SPFx origins before enabling the backend");
        for (String origin : allowedOrigins) {
            URI uri = URI.create(origin);
            boolean local =
                    "http".equals(uri.getScheme())
                            && ("localhost".equals(uri.getHost())
                                    || "127.0.0.1".equals(uri.getHost()));
            Assert.isTrue(
                    ("https".equals(uri.getScheme()) || local)
                            && uri.getHost() != null
                            && uri.getRawUserInfo() == null
                            && (uri.getRawPath() == null || uri.getRawPath().isEmpty())
                            && uri.getRawQuery() == null
                            && uri.getRawFragment() == null
                            && !origin.contains("*"),
                    "CORS origins must be exact HTTPS origins (or explicit localhost)");
        }
    }
}

package com.timematters.qstar.infrastructure.sharepoint;

import java.net.URI;
import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.util.Assert;

@Data
@ConfigurationProperties("qstar.sharepoint")
public class SharePointProperties {
    private String siteUrl = "";
    private String siteHostname = "";
    private String sitePath = "";
    private String issuesListName = "Q-Star Issues";
    private String progressListName = "Q-Star Progress Log";
    private String configListName = "Q-Star Config";
    private boolean betaAccessMode;

    public String getSiteUrl() {
        String value = siteUrl == null ? "" : siteUrl.strip();
        if (value.isEmpty() && siteHostname != null && !siteHostname.isBlank()) {
            value =
                    "https://"
                            + siteHostname.strip()
                            + (sitePath.startsWith("/") ? sitePath : "/" + sitePath);
        }
        return value.replaceAll("/+$", "");
    }

    public String tokenScope() {
        URI site = URI.create(getSiteUrl());
        return "https://" + site.getHost() + "/.default";
    }

    public void validate() {
        URI site = URI.create(getSiteUrl());
        String host = site.getHost();
        Assert.isTrue(
                "https".equals(site.getScheme())
                        && host != null
                        && (host.endsWith(".sharepoint.com")
                                || host.endsWith(".sharepoint.us")
                                || host.endsWith(".sharepoint.de")
                                || host.endsWith(".sharepoint.cn"))
                        && site.getPort() == -1
                        && site.getRawUserInfo() == null
                        && site.getRawQuery() == null
                        && site.getRawFragment() == null,
                "Configure one HTTPS SharePoint site URL before enabling the backend");
        Assert.hasText(issuesListName, "Issues list name is required");
        Assert.hasText(progressListName, "Progress list name is required");
        Assert.hasText(configListName, "Config list name is required");
    }
}

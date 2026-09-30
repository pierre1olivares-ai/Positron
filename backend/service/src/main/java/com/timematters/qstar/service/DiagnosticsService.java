package com.timematters.qstar.service;

import com.timematters.qstar.api.model.DiagnosticCheckATO;
import com.timematters.qstar.infrastructure.sharepoint.SharePointProperties;
import com.timematters.qstar.infrastructure.sharepoint.SharePointRestClient;
import com.timematters.qstar.model.DiagnosticCheck;
import com.timematters.qstar.model.mapper.DiagnosticCheckMapper;
import java.util.ArrayList;
import java.util.List;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

/**
 * Backend-to-SharePoint connection self-test. Complements ConnectionDiagnosticsService.ts on the
 * frontend, which only checks the browser-to-backend leg — this checks the signed-in
 * administrator's delegated access to the configured lists. It does not test writes, change consent
 * or establish other users' access.
 */
@Service
public class DiagnosticsService {

    private final SharePointProperties properties;
    private final SharePointRestClient sharePoint;
    private final DiagnosticCheckMapper mapper = new DiagnosticCheckMapper();

    @Autowired
    public DiagnosticsService(SharePointRestClient sharePoint, SharePointProperties properties) {
        this.sharePoint = sharePoint;
        this.properties = properties;
    }

    public List<DiagnosticCheckATO> runChecks() {
        List<DiagnosticCheck> results = new ArrayList<>();

        String siteId = null;
        try {
            siteId = sharePoint.getSiteId();
            results.add(
                    new DiagnosticCheck("Site access", "pass", "Resolved site id " + siteId + "."));
        } catch (Exception e) {
            results.add(new DiagnosticCheck("Site access", "fail", errorMessage(e)));
        }

        if (siteId != null) {
            checkList(results, "Q-Star Issues", properties.getIssuesListName());
            checkList(results, "Q-Star Progress Log", properties.getProgressListName());
            checkList(results, "Q-Star Config", properties.getConfigListName());
        } else {
            results.add(
                    new DiagnosticCheck(
                            "Lists",
                            "fail",
                            "Skipped — could not resolve the site, see Site access above."));
        }

        return results.stream().map(mapper::toATO).toList();
    }

    private void checkList(List<DiagnosticCheck> results, String label, String listName) {
        try {
            String listId = sharePoint.getListId(listName);
            results.add(
                    new DiagnosticCheck(
                            "List \"" + label + "\"", "pass", "Resolved list id " + listId + "."));
        } catch (Exception e) {
            results.add(new DiagnosticCheck("List \"" + label + "\"", "fail", errorMessage(e)));
        }
    }

    private String errorMessage(Exception e) {
        return "Delegated SharePoint access failed. Check site/list configuration and the caller's permissions.";
    }
}

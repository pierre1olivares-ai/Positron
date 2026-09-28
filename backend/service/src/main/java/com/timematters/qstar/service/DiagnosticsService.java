package com.timematters.qstar.service;

import com.timematters.qstar.api.model.DiagnosticCheckATO;
import com.timematters.qstar.infrastructure.sharepoint.SharePointGraphClient;
import com.timematters.qstar.model.DiagnosticCheck;
import com.timematters.qstar.model.mapper.DiagnosticCheckMapper;
import java.util.ArrayList;
import java.util.List;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

/**
 * Backend-to-SharePoint connection self-test. Complements
 * ConnectionDiagnosticsService.ts on the frontend, which only checks the browser-to-backend leg
 * — this checks whether the backend's own app identity (client-credentials) can actually reach
 * the configured SharePoint site and lists.
 */
@Service
public class DiagnosticsService {

    @Value("${qstar.sharepoint.issues-list}")
    private String issuesListName;

    @Value("${qstar.sharepoint.progress-list}")
    private String progressListName;

    private final SharePointGraphClient graphClient;
    private final DiagnosticCheckMapper mapper = new DiagnosticCheckMapper();

    @Autowired
    public DiagnosticsService(SharePointGraphClient graphClient) {
        this.graphClient = graphClient;
    }

    public List<DiagnosticCheckATO> runChecks() {
        List<DiagnosticCheck> results = new ArrayList<>();

        String siteId = null;
        try {
            siteId = graphClient.getSiteId();
            results.add(new DiagnosticCheck("Site access", "pass", "Resolved site id " + siteId + "."));
        } catch (Exception e) {
            results.add(new DiagnosticCheck("Site access", "fail", errorMessage(e)));
        }

        if (siteId != null) {
            checkList(results, "Q-Star Issues", issuesListName);
            checkList(results, "Q-Star Progress Log", progressListName);
            checkList(results, "Q-Star Config", "Q-Star Config");
        } else {
            results.add(
                    new DiagnosticCheck(
                            "Lists", "fail", "Skipped — could not resolve the site, see Site access above."));
        }

        return results.stream().map(mapper::toATO).toList();
    }

    private void checkList(List<DiagnosticCheck> results, String label, String listName) {
        try {
            String listId = graphClient.getListId(listName);
            results.add(new DiagnosticCheck("List \"" + label + "\"", "pass", "Resolved list id " + listId + "."));
        } catch (Exception e) {
            results.add(new DiagnosticCheck("List \"" + label + "\"", "fail", errorMessage(e)));
        }
    }

    private String errorMessage(Exception e) {
        return e.getMessage() != null ? e.getMessage() : e.getClass().getSimpleName();
    }
}

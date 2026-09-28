package com.timematters.qstar.api.controller;

import com.timematters.qstar.api.model.IssueATO;
import com.timematters.qstar.api.model.IssueCreateATO;
import com.timematters.qstar.api.model.IssuePatchATO;
import com.timematters.qstar.api.model.ProgressCreateATO;
import com.timematters.qstar.api.model.ProgressLogEntryATO;
import com.timematters.qstar.configuration.security.AuthorizationPolicy;
import com.timematters.qstar.service.IssueService;
import java.net.URI;
import java.util.List;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
@RequestMapping("${openapi.qStarIssueManager.base-path:/api/v1}")
public class IssuesApiController implements IssuesApi {
    private final IssueService issues;
    private final AuthorizationPolicy authorization;
    private final String basePath;

    public IssuesApiController(
            IssueService issues,
            AuthorizationPolicy authorization,
            @Value("${openapi.qStarIssueManager.base-path:/api/v1}") String basePath) {
        this.issues = issues;
        this.authorization = authorization;
        if (!basePath.startsWith("/") || basePath.startsWith("//"))
            throw new IllegalArgumentException(
                    "The API base path must be relative to this server.");
        this.basePath = basePath.replaceAll("/+$", "");
    }

    @Override
    public ResponseEntity<List<IssueATO>> getIssues() {
        return ResponseEntity.ok(issues.getAllIssues());
    }

    @Override
    public ResponseEntity<IssueATO> getIssue(Long id) {
        IssueATO issue = issues.getIssue(id);
        return new ResponseEntity<>(issue, versionHeaders(issue), HttpStatus.OK);
    }

    @Override
    public ResponseEntity<Object> createIssue(IssueCreateATO input) {
        authorization.requireManager();
        IssueATO issue = issues.createIssue(input);
        HttpHeaders headers = versionHeaders(issue);
        if (issue.getId() != null)
            headers.setLocation(URI.create(basePath + "/issues/" + issue.getId()));
        if (issue.getQsNumber() != null)
            headers.set("X-QStar-Reference", issue.getQsNumber().toString());
        return new ResponseEntity<>(issue, headers, HttpStatus.CREATED);
    }

    @Override
    public ResponseEntity<Object> updateIssue(Long id, String ifMatch, IssuePatchATO patch) {
        authorization.requireContributor();
        Object saved = issues.updateIssue(id, patch, ifMatch);
        HttpHeaders headers =
                saved instanceof IssueATO issue ? versionHeaders(issue) : new HttpHeaders();
        return new ResponseEntity<>(saved, headers, HttpStatus.OK);
    }

    @Override
    public ResponseEntity<Object> addProgressLogEntry(Long id, ProgressCreateATO input) {
        authorization.requireContributor();
        ProgressLogEntryATO entry = issues.addProgressLogEntry(id, input);
        HttpHeaders headers = new HttpHeaders();
        if (entry.getId() != null) {
            headers.set("X-QStar-Entry-Id", entry.getId().toString());
            headers.setLocation(
                    URI.create(basePath + "/issues/" + id + "/progress/" + entry.getId()));
        }
        return new ResponseEntity<>(entry, headers, HttpStatus.CREATED);
    }

    @Override
    public ResponseEntity<ProgressLogEntryATO> getProgressEntry(Long id, Long entryId) {
        return ResponseEntity.ok(issues.getProgressEntry(id, entryId));
    }

    private HttpHeaders versionHeaders(IssueATO issue) {
        var headers = new HttpHeaders();
        if (issue.getETag() != null && !issue.getETag().isBlank()) headers.setETag(issue.getETag());
        return headers;
    }
}

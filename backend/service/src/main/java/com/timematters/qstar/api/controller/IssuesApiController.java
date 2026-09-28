package com.timematters.qstar.api.controller;

import com.timematters.error.GenericError;
import com.timematters.qstar.api.model.IssueATO;
import com.timematters.qstar.api.model.IssueCreateATO;
import com.timematters.qstar.api.model.IssuePatchATO;
import com.timematters.qstar.api.model.ProgressLogEntryATO;
import com.timematters.qstar.service.IssueService;
import java.util.List;
import java.util.Optional;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.context.request.NativeWebRequest;

@Controller
@RequestMapping("${openapi.qStarIssueManager.base-path:/api/v1}")
public class IssuesApiController extends AbstractController implements IssuesApi {

    private final NativeWebRequest request;
    private final IssueService issueService;

    @Autowired
    public IssuesApiController(NativeWebRequest request, IssueService issueService) {
        this.request = request;
        this.issueService = issueService;
    }

    @Override
    public Optional<NativeWebRequest> getRequest() {
        return Optional.ofNullable(request);
    }

    @Override
    public ResponseEntity<List<IssueATO>> getIssues() throws GenericError {
        try {
            return new ResponseEntity<>(issueService.getAllIssues(), HttpStatus.OK);
        } catch (Exception exception) {
            rethrowAsUnexptedError(getRequest().orElse(null), exception);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @Override
    public ResponseEntity<IssueATO> createIssue(IssueCreateATO issueCreateATO) throws GenericError {
        try {
            return new ResponseEntity<>(issueService.createIssue(issueCreateATO), HttpStatus.CREATED);
        } catch (Exception exception) {
            rethrowAsUnexptedError(getRequest().orElse(null), exception);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @Override
    public ResponseEntity<IssueATO> updateIssue(Long id, IssuePatchATO issuePatchATO) throws GenericError {
        try {
            return new ResponseEntity<>(issueService.updateIssue(id, issuePatchATO), HttpStatus.OK);
        } catch (Exception exception) {
            rethrowAsUnexptedError(getRequest().orElse(null), exception);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @Override
    public ResponseEntity<ProgressLogEntryATO> addProgressLogEntry(Long id, ProgressLogEntryATO progressLogEntryATO)
            throws GenericError {
        try {
            return new ResponseEntity<>(issueService.addProgressLogEntry(id, progressLogEntryATO), HttpStatus.CREATED);
        } catch (Exception exception) {
            rethrowAsUnexptedError(getRequest().orElse(null), exception);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}

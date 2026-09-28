package com.timematters.qstar.api.controller;

import com.timematters.error.GenericError;
import com.timematters.qstar.api.model.DiagnosticCheckATO;
import com.timematters.qstar.configuration.security.AuthorizationPolicy;
import com.timematters.qstar.service.DiagnosticsService;
import java.util.List;
import java.util.Optional;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.context.request.NativeWebRequest;

@Controller
@RequestMapping("${openapi.qStarIssueManager.base-path:/api/v1}")
public class DiagnosticsApiController extends AbstractController implements DiagnosticsApi {

    private final NativeWebRequest request;
    private final DiagnosticsService diagnosticsService;
    private final AuthorizationPolicy authorization;

    @Autowired
    public DiagnosticsApiController(
            NativeWebRequest request,
            DiagnosticsService diagnosticsService,
            AuthorizationPolicy authorization) {
        this.request = request;
        this.diagnosticsService = diagnosticsService;
        this.authorization = authorization;
    }

    @Override
    public Optional<NativeWebRequest> getRequest() {
        return Optional.ofNullable(request);
    }

    @Override
    public ResponseEntity<List<DiagnosticCheckATO>> getDiagnostics() throws GenericError {
        authorization.requireAdmin();
        return ResponseEntity.ok(diagnosticsService.runChecks());
    }
}

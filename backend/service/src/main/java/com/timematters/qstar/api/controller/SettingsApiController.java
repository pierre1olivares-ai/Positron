package com.timematters.qstar.api.controller;

import com.timematters.error.GenericError;
import com.timematters.qstar.api.model.SettingsATO;
import com.timematters.qstar.configuration.security.AuthorizationPolicy;
import com.timematters.qstar.infrastructure.sharepoint.SettingsRepository;
import com.timematters.qstar.model.mapper.SettingsMapper;
import java.util.Optional;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.context.request.NativeWebRequest;

@Controller
@RequestMapping("${openapi.qStarIssueManager.base-path:/api/v1}")
public class SettingsApiController extends AbstractController implements SettingsApi {

    private final NativeWebRequest request;
    private final SettingsRepository settingsRepository;
    private final AuthorizationPolicy authorization;
    private final SettingsMapper settingsMapper = new SettingsMapper();

    @Autowired
    public SettingsApiController(
            NativeWebRequest request,
            SettingsRepository settingsRepository,
            AuthorizationPolicy authorization) {
        this.request = request;
        this.settingsRepository = settingsRepository;
        this.authorization = authorization;
    }

    @Override
    public Optional<NativeWebRequest> getRequest() {
        return Optional.ofNullable(request);
    }

    @Override
    public ResponseEntity<SettingsATO> getSettings() throws GenericError {
        return ResponseEntity.ok(settingsMapper.toATO(settingsRepository.load()));
    }

    @Override
    public ResponseEntity<SettingsATO> saveSettings(SettingsATO settingsATO) throws GenericError {
        authorization.requireAdmin();
        var saved = settingsRepository.save(settingsMapper.fromATO(settingsATO));
        return ResponseEntity.ok(settingsMapper.toATO(saved));
    }
}

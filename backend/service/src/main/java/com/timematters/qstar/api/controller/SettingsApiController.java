package com.timematters.qstar.api.controller;

import com.timematters.error.GenericError;
import com.timematters.qstar.api.model.SettingsATO;
import com.timematters.qstar.infrastructure.sharepoint.SettingsRepository;
import com.timematters.qstar.model.mapper.SettingsMapper;
import java.util.Optional;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.context.request.NativeWebRequest;

@Controller
@RequestMapping("${openapi.qStarIssueManager.base-path:/api/v1}")
public class SettingsApiController extends AbstractController implements SettingsApi {

    private final NativeWebRequest request;
    private final SettingsRepository settingsRepository;
    private final SettingsMapper settingsMapper = new SettingsMapper();

    @Autowired
    public SettingsApiController(NativeWebRequest request, SettingsRepository settingsRepository) {
        this.request = request;
        this.settingsRepository = settingsRepository;
    }

    @Override
    public Optional<NativeWebRequest> getRequest() {
        return Optional.ofNullable(request);
    }

    @Override
    public ResponseEntity<SettingsATO> getSettings() throws GenericError {
        try {
            return new ResponseEntity<>(settingsMapper.toATO(settingsRepository.load()), HttpStatus.OK);
        } catch (Exception exception) {
            rethrowAsUnexptedError(getRequest().orElse(null), exception);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @Override
    public ResponseEntity<SettingsATO> saveSettings(SettingsATO settingsATO) throws GenericError {
        try {
            var saved = settingsRepository.save(settingsMapper.fromATO(settingsATO));
            return new ResponseEntity<>(settingsMapper.toATO(saved), HttpStatus.OK);
        } catch (Exception exception) {
            rethrowAsUnexptedError(getRequest().orElse(null), exception);
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}

package com.timematters.qstar.api;

import com.timematters.error.ErrorCode;
import com.timematters.qstar.api.model.ErrorATO;
import jakarta.servlet.http.HttpServletRequest;
import java.util.UUID;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.boot.web.servlet.error.ErrorController;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

// NOTE: the template's version of this class extended
// org.springframework.boot.autoconfigure.web.servlet.error.AbstractErrorController and overrode
// getErrorPath(). Neither exists in Spring Boot 3.3.4: ErrorController is now an empty marker
// interface (no getErrorPath(), no required constructor args) — the error path is wired purely
// through the @RequestMapping below plus server.error.path/error.path properties (both default
// to /error, matching ERROR_PATH).
@RestController
@RequestMapping({CustomErrorController.ERROR_PATH})
public class CustomErrorController implements ErrorController {
    private static final Logger log = LogManager.getLogger();

    static final String ERROR_PATH = "/error";

    @RequestMapping
    public ResponseEntity<ErrorATO> error(HttpServletRequest request) {
        log.error("Unexpected error with following request: {}", request);

        ErrorATO errorResponseATO = new ErrorATO();
        errorResponseATO.setStatus(500);
        errorResponseATO.setCode(ErrorCode.DEFAULT_ERROR.toString());
        errorResponseATO.setTitle(ErrorCode.DEFAULT_ERROR.name());
        errorResponseATO.setDetails(
                "Please contact the developers and give them the ID of this Error.");
        errorResponseATO.setId(UUID.randomUUID().toString());
        return new ResponseEntity<>(errorResponseATO, HttpStatus.INTERNAL_SERVER_ERROR);
    }
}

package com.timematters.qstar.api.controller;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.timematters.error.ErrorCode;
import com.timematters.error.GenericError;
import com.timematters.qstar.api.CustomOffsetDateTimeMapper;
import java.util.UUID;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.web.context.request.NativeWebRequest;

public class AbstractController {
    private static final Logger log = LogManager.getLogger();

    protected ObjectMapper responseMapper;

    public AbstractController() {
        this.responseMapper = new CustomOffsetDateTimeMapper();
    }

    protected void rethrowAsUnexptedError(NativeWebRequest request, Exception exception)
            throws GenericError {
        log.error("Unexpected error: {} with following request: {}", exception, request);

        GenericError unexpectedError = new GenericError();
        unexpectedError.setStatus(500);
        unexpectedError.setCode(ErrorCode.DEFAULT_ERROR.toString());
        unexpectedError.setTitle(ErrorCode.DEFAULT_ERROR.name());
        unexpectedError.setDetails(
                "Please contact the developers and give them the ID of this Error.");
        unexpectedError.setId(UUID.randomUUID().toString());
        throw unexpectedError;
    }
}

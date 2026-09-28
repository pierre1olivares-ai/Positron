package com.timematters.error;

import com.timematters.qstar.api.model.ErrorATO;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import org.springframework.http.HttpStatus;

public class ValidationError extends GenericError {
    private static final long serialVersionUID = 8437603878870646233L;

    private List<ValidationViolation> validationViolations;

    public ValidationError(ErrorCode errorCode, ValidationViolations validationViolations) {
        super();
        this.validationViolations = validationViolations.getValidationViolations();
        this.setId(UUID.randomUUID().toString());
        this.setStatus(HttpStatus.BAD_REQUEST.value());
        this.setCode(errorCode.name());
        this.setTitle(errorCode.toString());
    }

    @Override
    public ErrorATO toErrorResponse() {
        ErrorATO errorResponse = super.toErrorResponse();
        Map<String, Object> violationsContainer = new HashMap<>();
        violationsContainer.put("validationErrors", validationViolations);
        errorResponse.setMeta(violationsContainer);
        return errorResponse;
    }
}

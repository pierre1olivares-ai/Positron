package com.timematters.qstar.api;

import com.timematters.error.ErrorCode;
import com.timematters.error.GenericError;
import com.timematters.error.ValidationError;
import com.timematters.error.ValidationViolations;
import com.timematters.qstar.api.model.ErrorATO;
import com.timematters.qstar.service.IssueOperationException;
import jakarta.validation.ConstraintViolation;
import jakarta.validation.ConstraintViolationException;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.UUID;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.hibernate.validator.internal.engine.path.PathImpl;
import org.springframework.beans.TypeMismatchException;
import org.springframework.core.Ordered;
import org.springframework.core.annotation.Order;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.HttpStatusCode;
import org.springframework.http.ResponseEntity;
import org.springframework.http.converter.HttpMessageNotReadableException;
import org.springframework.security.access.AccessDeniedException;
import org.springframework.validation.FieldError;
import org.springframework.validation.ObjectError;
import org.springframework.web.HttpMediaTypeNotSupportedException;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.MissingServletRequestParameterException;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.client.HttpStatusCodeException;
import org.springframework.web.client.RestClientException;
import org.springframework.web.context.request.WebRequest;
import org.springframework.web.server.ResponseStatusException;
import org.springframework.web.servlet.NoHandlerFoundException;
import org.springframework.web.servlet.mvc.method.annotation.ResponseEntityExceptionHandler;

@Order(Ordered.HIGHEST_PRECEDENCE)
@ControllerAdvice
public class CustomExceptionHandler extends ResponseEntityExceptionHandler {
    private static final Logger log = LogManager.getLogger();

    @Override
    protected ResponseEntity<Object> handleExceptionInternal(
            Exception exception,
            Object body,
            HttpHeaders headers,
            HttpStatusCode status,
            WebRequest request) {
        return new ResponseEntity<>(
                Map.of(
                        "status",
                        status.value(),
                        "code",
                        "REQUEST_FAILED",
                        "details",
                        "The request could not be processed"),
                headers,
                status);
    }

    @ExceptionHandler(AccessDeniedException.class)
    public ResponseEntity<?> handleAccessDenied(AccessDeniedException exception) {
        return ResponseEntity.status(403)
                .body(
                        Map.of(
                                "status",
                                403,
                                "code",
                                "FORBIDDEN",
                                "details",
                                "This operation is not permitted for your Q-Star role or ownership"));
    }

    @ExceptionHandler(IssueOperationException.class)
    public ResponseEntity<?> handleIssueOperation(IssueOperationException exception) {
        return ResponseEntity.status(exception.status())
                .body(
                        Map.of(
                                "status",
                                exception.status(),
                                "code",
                                exception.code(),
                                "details",
                                exception.getMessage(),
                                "meta",
                                exception.meta()));
    }

    @ExceptionHandler(ResponseStatusException.class)
    public ResponseEntity<?> handleStatus(ResponseStatusException exception) {
        return ResponseEntity.status(exception.getStatusCode())
                .body(
                        Map.of(
                                "status",
                                exception.getStatusCode().value(),
                                "code",
                                "REQUEST_FAILED",
                                "details",
                                exception.getReason() == null
                                        ? "The request could not be processed"
                                        : exception.getReason()));
    }

    @ExceptionHandler(HttpStatusCodeException.class)
    public ResponseEntity<?> handleSharePointStatus(HttpStatusCodeException exception) {
        int downstream = exception.getStatusCode().value();
        int status = List.of(401, 403, 404, 412).contains(downstream) ? downstream : 502;
        String message =
                status == 403
                        ? "SharePoint denied the signed-in user's access"
                        : status == 404
                                ? "The requested SharePoint resource was not found"
                                : status == 412
                                        ? "The item changed; reload before retrying"
                                        : "Delegated SharePoint access failed";
        return ResponseEntity.status(status)
                .body(
                        Map.of(
                                "status",
                                status,
                                "code",
                                "SHAREPOINT_REQUEST_FAILED",
                                "details",
                                message));
    }

    @ExceptionHandler(RestClientException.class)
    public ResponseEntity<?> handleSharePointTransport(RestClientException exception) {
        return ResponseEntity.status(502)
                .body(
                        Map.of(
                                "status",
                                502,
                                "code",
                                "SHAREPOINT_UNAVAILABLE",
                                "details",
                                "SharePoint could not be reached; reconcile pending writes before retrying"));
    }

    /** Handles our internal Error and builds a response out of it. */
    @ExceptionHandler(GenericError.class)
    private ResponseEntity<Object> handleError(GenericError error) {
        return buildResponseEntity(error.toErrorResponse());
    }

    @Override
    protected ResponseEntity<Object> handleNoHandlerFoundException(
            NoHandlerFoundException ex,
            HttpHeaders headers,
            HttpStatusCode status,
            WebRequest request) {
        ErrorATO errorResponseATO = new ErrorATO();
        errorResponseATO.setStatus(404);
        errorResponseATO.setCode(ErrorCode.RESOURCE_NOT_FOUND.toString());
        errorResponseATO.setTitle(ErrorCode.RESOURCE_NOT_FOUND.name());
        errorResponseATO.setId(UUID.randomUUID().toString());
        return buildResponseEntity(errorResponseATO);
    }

    @Override
    protected ResponseEntity<Object> handleTypeMismatch(
            TypeMismatchException ex,
            HttpHeaders headers,
            HttpStatusCode status,
            WebRequest request) {
        return ResponseEntity.badRequest()
                .body(
                        Map.of(
                                "status",
                                400,
                                "code",
                                "INVALID_PARAMETER",
                                "details",
                                "A request parameter has an invalid value"));
    }

    /* ----------------------------------------
     * Validation Errors
     * ---------------------------------------- */

    /** Handles jakarta.validation.ConstraintViolationException. Thrown when @Validated fails. */
    @ExceptionHandler(ConstraintViolationException.class)
    protected ResponseEntity<Object> handleContraintViolationException(
            ConstraintViolationException exception) {
        ValidationViolations validationViolations = new ValidationViolations();
        Set<ConstraintViolation<?>> constraintViolations = exception.getConstraintViolations();

        for (ConstraintViolation<?> constraintViolation : constraintViolations) {
            validationViolations.add(
                    ((PathImpl) constraintViolation.getPropertyPath()).getLeafNode().asString(),
                    constraintViolation.getMessage(),
                    constraintViolation.getInvalidValue());
        }
        ValidationError validationError =
                new ValidationError(ErrorCode.BODY_NOT_VALID, validationViolations);
        return buildResponseEntity(validationError.toErrorResponse());
    }

    /** Handle HttpMessageNotReadableException. Happens when request JSON is malformed. */
    @Override
    protected ResponseEntity<Object> handleHttpMessageNotReadable(
            HttpMessageNotReadableException exception,
            HttpHeaders headers,
            HttpStatusCode status,
            WebRequest request) {
        ErrorATO errorResponseATO = new ErrorATO();
        errorResponseATO.setStatus(HttpStatus.BAD_REQUEST.value());
        errorResponseATO.setCode(ErrorCode.BODY_NOT_VALID.toString());
        errorResponseATO.setTitle(ErrorCode.BODY_NOT_VALID.name());
        errorResponseATO.setDetails("The request body is malformed or contains an invalid value");
        errorResponseATO.setId(UUID.randomUUID().toString());
        return buildResponseEntity(errorResponseATO);
    }

    /** Handle MethodArgumentNotValidException. Triggered when an object fails @Valid validation. */
    @Override
    protected ResponseEntity<Object> handleMethodArgumentNotValid(
            MethodArgumentNotValidException exception,
            HttpHeaders headers,
            HttpStatusCode status,
            WebRequest request) {
        ValidationViolations validationViolations = new ValidationViolations();
        List<FieldError> fieldErrors = exception.getBindingResult().getFieldErrors();
        for (FieldError fieldError : fieldErrors) {
            validationViolations.add(
                    fieldError.getField(), fieldError.getCode(), fieldError.getRejectedValue());
        }
        List<ObjectError> objectErrors = exception.getBindingResult().getGlobalErrors();
        for (ObjectError objectError : objectErrors) {
            validationViolations.add(objectError.getObjectName(), objectError.getCode(), "");
        }
        ValidationError validationError =
                new ValidationError(ErrorCode.BODY_NOT_VALID, validationViolations);
        return buildResponseEntity(validationError.toErrorResponse());
    }

    /**
     * Handle MissingServletRequestParameterException. Triggered when a 'required' request parameter
     * is missing.
     */
    @Override
    protected ResponseEntity<Object> handleMissingServletRequestParameter(
            MissingServletRequestParameterException exception,
            HttpHeaders headers,
            HttpStatusCode status,
            WebRequest request) {
        ErrorATO errorResponseATO = new ErrorATO();
        errorResponseATO.setStatus(HttpStatus.BAD_REQUEST.value());
        errorResponseATO.setCode(ErrorCode.QUERY_PARAMETER_MISSING.toString());
        errorResponseATO.setTitle(ErrorCode.QUERY_PARAMETER_MISSING.name());
        errorResponseATO.setDetails(exception.getMessage());
        errorResponseATO.setId(UUID.randomUUID().toString());
        return buildResponseEntity(errorResponseATO);
    }

    /**
     * Handle HttpMediaTypeNotSupportedException. This one triggers when JSON is invalid as well.
     */
    @Override
    protected ResponseEntity<Object> handleHttpMediaTypeNotSupported(
            HttpMediaTypeNotSupportedException ex,
            HttpHeaders headers,
            HttpStatusCode status,
            WebRequest request) {
        ErrorATO errorResponseATO = new ErrorATO();
        errorResponseATO.setStatus(HttpStatus.BAD_REQUEST.value());
        errorResponseATO.setCode(ErrorCode.BODY_NOT_VALID.toString());
        errorResponseATO.setTitle(ErrorCode.BODY_NOT_VALID.name());
        errorResponseATO.setId(UUID.randomUUID().toString());
        return buildResponseEntity(errorResponseATO);
    }

    /* ----------------------------------------
     * Helpers
     * ---------------------------------------- */
    private ResponseEntity<Object> buildResponseEntity(ErrorATO errorResponseATO) {
        HttpStatus httpStatus = HttpStatus.INTERNAL_SERVER_ERROR;
        if (errorResponseATO.getStatus() != null) {
            httpStatus = HttpStatus.valueOf(errorResponseATO.getStatus());
        }
        return new ResponseEntity<>(errorResponseATO, httpStatus);
    }

    private ErrorATO createStandardErrorResponse() {
        ErrorATO errorResponseATO = new ErrorATO();
        errorResponseATO.setStatus(500);
        errorResponseATO.setCode(ErrorCode.DEFAULT_ERROR.toString());
        errorResponseATO.setTitle(ErrorCode.DEFAULT_ERROR.name());
        errorResponseATO.setDetails(
                "Please contact the developers and give them the ID of this Error.");
        errorResponseATO.setId(UUID.randomUUID().toString());
        return errorResponseATO;
    }
}

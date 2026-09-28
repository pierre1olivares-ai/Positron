package com.timematters.error;

import java.io.Serializable;

// Previously extended a generated ValidationErrorATO for its source/error/value fields. That
// type no longer exists — see api-contract/contract.yaml's Error schema: validation details now
// travel in Error's free-form `meta` object (as ValidationError.toErrorResponse() already did,
// {"validationErrors": [...]}) rather than as their own response shape, so this only needs to be
// a plain, independently serializable POJO now.
public class ValidationViolation implements Serializable {
    private static final long serialVersionUID = -54385981982955366L;

    private String source;
    private String error;
    private String value;

    public ValidationViolation(String source, String error, String value) {
        this.source = source;
        this.error = error;
        this.value = value;
    }

    public String getSource() {
        return source;
    }

    public void setSource(String source) {
        this.source = source;
    }

    public String getError() {
        return error;
    }

    public void setError(String error) {
        this.error = error;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }
}

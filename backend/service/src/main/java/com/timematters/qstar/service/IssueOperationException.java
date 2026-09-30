package com.timematters.qstar.service;

import java.util.Map;

/**
 * Expected application failures, preserved by the HTTP boundary instead of becoming generic 500s.
 */
public class IssueOperationException extends RuntimeException {
    private final int status;
    private final String code;
    private final Map<String, Object> meta;

    public IssueOperationException(int status, String code, String message) {
        this(status, code, message, Map.of());
    }

    public IssueOperationException(
            int status, String code, String message, Map<String, Object> meta) {
        super(message);
        this.status = status;
        this.code = code;
        this.meta = meta;
    }

    public int status() {
        return status;
    }

    public String code() {
        return code;
    }

    public Map<String, Object> meta() {
        return meta;
    }
}

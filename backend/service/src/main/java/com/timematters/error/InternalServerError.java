package com.timematters.error;

import java.util.UUID;
import org.springframework.http.HttpStatus;

public class InternalServerError extends GenericError {
    private static final long serialVersionUID = 8685348701929964427L;

    public InternalServerError(String errorMessage) {
        super();
        this.setId(UUID.randomUUID().toString());
        this.setCode("internal_server_error");
        this.setStatus(HttpStatus.INTERNAL_SERVER_ERROR.value());
        this.setTitle("Internal Server Error");
        this.setDetails(errorMessage);
    }
}

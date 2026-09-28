package com.timematters.error;

public enum ErrorCode {
    DEFAULT_ERROR("Oh no we messed up. This should not happen."),
    RESOURCE_NOT_FOUND("This ressource is not available"),
    QUERY_PARAMETER_MISSING("You need to provide a query parameter"),
    BODY_NOT_VALID("The body you provided is not valid"),
    UPSTREAM_ERROR("The call to SharePoint/Graph failed");

    private final String string;

    ErrorCode(final String string) {
        this.string = string;
    }

    @Override
    public String toString() {
        return string;
    }
}

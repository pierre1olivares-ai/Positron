package com.timematters.qstar.infrastructure.sharepoint;

/**
 * A successful write response could not supply its identity. It must never be retried
 * automatically.
 */
public class AcceptedWriteException extends RuntimeException {
    public AcceptedWriteException() {
        super(
                "The write was accepted. Reload before submitting again; its response could not be read.");
    }
}

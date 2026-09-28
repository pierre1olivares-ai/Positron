package com.timematters.qstar.api.authentication.user;

import java.util.UUID;

public class UserId {
    private UUID userId;

    public UserId(UUID userId) {
        this.userId = userId;
    }

    public UserId(String userId) {
        this.userId = UUID.fromString(userId);
    }

    public UUID getUserId() {
        return userId;
    }

    public void setUserId(UUID userId) {
        this.userId = userId;
    }

    @Override
    public String toString() {
        return this.userId.toString();
    }

    public static UserId fromString(String userId) {
        assert userId.isEmpty() == false;
        return new UserId(userId);
    }
}

package com.timematters.qstar.configuration.security;

public interface CurrentUserProvider {
    CallerContext currentUser();
}

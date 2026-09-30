package com.timematters.qstar.configuration.security;

/** Server-derived identity. SharePoint numeric identity is used for issue ownership. */
public record CallerContext(String role, long sharePointUserId, String displayName, String email) {}

package com.timematters.qstar.configuration.security;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;

@Data
@ConfigurationProperties("qstar.backend")
public class BackendProperties {
    private boolean enabled;
}

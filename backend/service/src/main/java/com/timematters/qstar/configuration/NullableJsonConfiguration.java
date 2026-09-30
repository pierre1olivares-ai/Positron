package com.timematters.qstar.configuration;

import java.time.Clock;
import java.time.ZoneId;
import org.openapitools.jackson.nullable.JsonNullableModule;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class NullableJsonConfiguration {
    @Bean
    public JsonNullableModule jsonNullableModule() {
        return new JsonNullableModule();
    }

    @Bean
    public Clock qstarClock(@Value("${qstar.calendar-zone:Europe/Amsterdam}") String zone) {
        return Clock.system(ZoneId.of(zone));
    }
}

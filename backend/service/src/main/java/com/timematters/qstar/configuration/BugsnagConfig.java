package com.timematters.qstar.configuration;

import com.bugsnag.Bugsnag;
import com.bugsnag.BugsnagSpringConfiguration;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;
import org.springframework.boot.info.BuildProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Import;

@Configuration
@ConditionalOnProperty(name = "qstar.telemetry.bugsnag.enabled", havingValue = "true")
@Import(BugsnagSpringConfiguration.class)
public class BugsnagConfig {
    private static final String STAGE_DEVELOP = "develop";
    private static final String STAGE_INTEGRATION = "integration";
    private static final String STAGE_PRODUCTION = "production";

    @Autowired BuildProperties buildProperties;

    @Value("${tm.stage}")
    String stage;

    @Value("${tm.qstar.bugsnag.api.key}")
    String apiKey;

    @Bean
    public Bugsnag bugsnag() {
        Bugsnag bugsnag = new Bugsnag(apiKey);
        bugsnag.setAppVersion(buildProperties.getVersion());
        if (stage == null) {
            stage = "local";
        }
        bugsnag.setReleaseStage(stage);
        bugsnag.setNotifyReleaseStages(STAGE_DEVELOP, STAGE_INTEGRATION, STAGE_PRODUCTION);
        return bugsnag;
    }
}

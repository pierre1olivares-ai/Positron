package com.timematters.qstar;

import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.ExitCodeGenerator;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.autoconfigure.domain.EntityScan;
import org.springframework.boot.autoconfigure.flyway.FlywayAutoConfiguration;
import org.springframework.boot.autoconfigure.jdbc.DataSourceAutoConfiguration;
import org.springframework.context.annotation.ComponentScan;

// NOTE: the template's Application.java shipped with a second, inner
// `WebSecurityConfig` class extending the legacy `WebSecurityConfigurerAdapter`
// (removed in Spring Security 6 / incompatible with the Spring Boot 3.3.4
// declared in build.gradle) plus a wildcard-origin CORS `WebMvcConfigurer`
// bean. Both were dead/broken code, superseded by the real, working config in
// configuration/security/WebSecurityConfig.java (which this class's
// @ComponentScan of "com.timematters.qstar" already picks up as a subpackage).
// Removed rather than fixed in place, since keeping two competing security
// configs around is worse than one correct one.
// Q-Star stores its data in SharePoint. Keep the database template as unwired scaffolding.
@SpringBootApplication(exclude = {DataSourceAutoConfiguration.class, FlywayAutoConfiguration.class})
@ComponentScan(
        basePackages = {
            "com.timematters.qstar",
            "com.timematters.qstar.api",
            "com.timematters.qstar.api.controller",
            "com.timematters.qstar.api.model",
            "com.timematters.qstar.service.model",
            "org.openapitools"
        })
@EntityScan(basePackages = {"com.timematters.qstar.persistence.entity"})
public class Application implements CommandLineRunner {

    @Override
    public void run(String... arg0) throws Exception {
        if (arg0.length > 0 && arg0[0].equals("exitcode")) {
            throw new ExitException();
        }
    }

    public static void main(String[] args) {
        new SpringApplication(Application.class).run(args);
    }

    static class ExitException extends RuntimeException implements ExitCodeGenerator {
        private static final long serialVersionUID = 1L;

        @Override
        public int getExitCode() {
            return 10;
        }
    }
}

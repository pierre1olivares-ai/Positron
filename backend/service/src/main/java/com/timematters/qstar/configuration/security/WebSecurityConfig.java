package com.timematters.qstar.configuration.security;

import com.azure.spring.cloud.autoconfigure.implementation.aad.security.AadResourceServerHttpSecurityConfigurer;
import com.timematters.qstar.configuration.StageType;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import lombok.RequiredArgsConstructor;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.security.config.Customizer;
import org.springframework.security.config.annotation.method.configuration.EnableMethodSecurity;
import org.springframework.security.config.annotation.web.builders.HttpSecurity;
import org.springframework.security.config.annotation.web.configuration.EnableWebSecurity;
import org.springframework.security.web.SecurityFilterChain;
import org.springframework.web.cors.CorsConfiguration;
import org.springframework.web.cors.CorsConfigurationSource;
import org.springframework.web.cors.UrlBasedCorsConfigurationSource;

// NOTE: the template's version of this class extended
// AadResourceServerWebSecurityConfigurerAdapter, which doesn't exist in
// com.azure.spring:spring-cloud-azure-starter-active-directory:5.17.1 (the version build.gradle
// actually declares) — confirmed by inspecting the real jar; that whole *WebSecurityConfigurerAdapter
// inheritance style was removed in Spring Security 6 anyway (Spring Boot 3). The current
// equivalent is AadResourceServerHttpSecurityConfigurer, an AbstractHttpConfigurer applied via
// HttpSecurity.with(...) inside an ordinary @Bean SecurityFilterChain method — the standard
// Spring Security 6 component style, replacing the old subclass-and-override approach below.
// Likewise @EnableGlobalMethodSecurity is deprecated in favor of @EnableMethodSecurity.
@EnableWebSecurity
@EnableMethodSecurity(prePostEnabled = true)
@RequiredArgsConstructor
public class WebSecurityConfig {

    @Value("${tm.stage}")
    private String stage;

    @Bean
    public SecurityFilterChain filterChain(HttpSecurity http) throws Exception {
        http.with(AadResourceServerHttpSecurityConfigurer.aadResourceServer(), Customizer.withDefaults());
        // Enable CORS with the configuration below
        http.cors(Customizer.withDefaults());
        // Disable CSRF check on each POST/PUT/DELETE — this is a stateless bearer-token API, not
        // a cookie-session one, so CSRF protection doesn't apply.
        http.csrf(csrf -> csrf.disable());
        // any request must be an authorized request
        http.authorizeHttpRequests(authorize -> authorize.anyRequest().authenticated());
        return http.build();
    }

    @Bean
    public CorsConfigurationSource corsConfigurationSource() {
        CorsConfiguration configuration = new CorsConfiguration();
        List<String> corsLocations = new ArrayList<>();
        corsLocations.add("https://qstar*.time-matters.com*");
        if (!StageType.fromValue(stage).equals(StageType.STAGE_PRODUCTION)) {
            corsLocations.add("http://localhost:4200*");
        }
        configuration.setAllowedOriginPatterns(corsLocations);
        configuration.setAllowedMethods(Arrays.asList("*"));
        configuration.setAllowedHeaders(Arrays.asList("*"));
        configuration.setAllowCredentials(true);
        UrlBasedCorsConfigurationSource source = new UrlBasedCorsConfigurationSource();
        source.registerCorsConfiguration("/**", configuration);
        return source;
    }
}

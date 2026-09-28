package com.timematters.qstar.configuration.security;

import static org.junit.jupiter.api.Assertions.*;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.*;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.*;

import com.timematters.qstar.api.authentication.MicrosoftGraphClient;
import com.timematters.qstar.infrastructure.QstarDatabase;
import javax.sql.DataSource;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.context.ApplicationContext;
import org.springframework.security.oauth2.client.registration.ClientRegistrationRepository;
import org.springframework.test.web.servlet.MockMvc;

@SpringBootTest(properties = "qstar.backend.enabled=false")
@AutoConfigureMockMvc
class DisabledBackendTest {
    @Autowired MockMvc mvc;
    @Autowired ApplicationContext context;

    @Test
    void bootsWithoutSecretsAndCannotExposeOrMutateData() throws Exception {
        mvc.perform(get("/api/v1/me"))
                .andExpect(status().isServiceUnavailable())
                .andExpect(jsonPath("$.code").value("BACKEND_DISABLED"));
        mvc.perform(post("/api/v1/issues").contentType("application/json").content("{}"))
                .andExpect(status().isServiceUnavailable());
        assertTrue(context.getBeansOfType(MicrosoftGraphClient.class).isEmpty());
        assertTrue(context.getBeansOfType(QstarDatabase.class).isEmpty());
        assertTrue(context.getBeansOfType(DataSource.class).isEmpty());
        assertTrue(context.getBeansOfType(ClientRegistrationRepository.class).isEmpty());
        assertTrue(context.getBeansOfType(com.bugsnag.Bugsnag.class).isEmpty());
    }
}

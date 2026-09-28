package com.timematters.qstar.infrastructure.integration;

import static org.junit.jupiter.api.parallel.ExecutionMode.SAME_THREAD;

import com.timematters.qstar.helper.Constants;
import com.timematters.qstar.helper.ModelATOHelper;
import com.timematters.qstar.helper.ModelDTOHelper;
import com.timematters.qstar.helper.ModelHelper;
import io.zonky.test.db.AutoConfigureEmbeddedDatabase;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.SQLException;
import javax.sql.DataSource;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.parallel.Execution;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.ActiveProfiles;
import org.springframework.test.context.junit.jupiter.SpringExtension;
import org.springframework.test.web.client.MockRestServiceServer;
import org.springframework.web.client.RestTemplate;

// NOTE: the template's version of this class had a @MockBean on
// infrastructure.sample.SampleRepository (the placeholder example's JDBI
// repository), removed along with the rest of the Sample/ComplexSample
// domain. Since IssueRepository talks to Microsoft Graph over HTTP rather
// than JDBI, tests against it should use createMockRestServer(...) below
// (MockRestServiceServer bound to the injected RestTemplate) instead of a
// @MockBean repository.
@ExtendWith(SpringExtension.class)
@AutoConfigureEmbeddedDatabase(beanName = "dataSource")
@SpringBootTest()
@Execution(SAME_THREAD)
@ActiveProfiles("integration")
@Tag("slow")
public class AbstractIntegrationTest extends Constants {

    @Autowired
    @Qualifier("customRestTemplate")
    RestTemplate platformRestTemplate;

    @Autowired RestTemplate restTemplate;

    @Autowired DataSource dataSource;

    protected ModelDTOHelper modelDTOHelper;
    protected ModelHelper modelHelper;
    protected ModelATOHelper modelATOHelper;

    public AbstractIntegrationTest() {
        this.modelHelper = new ModelHelper();
        this.modelATOHelper = new ModelATOHelper();
        this.modelDTOHelper = new ModelDTOHelper();
    }

    protected void clearTable(String tableName) throws SQLException {
        dataSource.getConnection().createStatement().execute("DELETE FROM " + tableName + ";");
    }

    protected MockRestServiceServer createMockRestServer(RestTemplate restTemplate) {
        return MockRestServiceServer.bindTo(restTemplate).ignoreExpectOrder(true).build();
    }

    protected void prefillDatabase(String filename) throws Exception {
        String sqlStatement = readFileFromResources("sql/" + filename);
        dataSource.getConnection().createStatement().execute(sqlStatement);
    }

    protected String readFileFromResources(String filename) throws Exception {
        URL resource = getClass().getClassLoader().getResource(filename);
        byte[] bytes = Files.readAllBytes(Paths.get(resource.toURI()));
        return new String(bytes);
    }
}

package com.timematters.qstar.infrastructure;

import com.timematters.qstar.api.CustomOffsetDateTimeMapper;
import com.timematters.qstar.helper.ModelDTOHelper;
import com.timematters.qstar.helper.ModelHelper;
import io.zonky.test.db.postgres.embedded.EmbeddedPostgres;
import java.io.IOException;
import java.sql.SQLException;
import java.sql.Statement;
import javax.sql.DataSource;
import org.flywaydb.core.Flyway;
import org.jdbi.v3.core.Jdbi;
import org.jdbi.v3.jackson2.Jackson2Config;
import org.jdbi.v3.jackson2.Jackson2Plugin;
import org.jdbi.v3.json.JsonPlugin;
import org.jdbi.v3.postgres.PostgresPlugin;
import org.jdbi.v3.sqlobject.SqlObjectPlugin;

public class AbstractRepositoryTest {
    protected ModelDTOHelper modelDTOHelper;
    protected ModelHelper modelHelper;
    protected EmbeddedPostgres embeddedPostgres;
    protected Jdbi jdbiInstance;

    public AbstractRepositoryTest() throws IOException {
        prepareDatabase();
        modelDTOHelper = new ModelDTOHelper();
        modelHelper = new ModelHelper();
    }

    protected void prepareDatabase() throws IOException {
        embeddedPostgres = EmbeddedPostgres.start();
        DataSource dataSource = embeddedPostgres.getPostgresDatabase();
        migrateDatabase(dataSource);
        jdbiInstance =
                Jdbi.create(dataSource)
                        .installPlugin(new PostgresPlugin())
                        .installPlugin(new SqlObjectPlugin())
                        .installPlugin(new JsonPlugin())
                        .installPlugin(new Jackson2Plugin());

        jdbiInstance.getConfig(Jackson2Config.class).setMapper(new CustomOffsetDateTimeMapper());
    }

    protected void clearTable(String tableName) throws SQLException {
        createStatement().execute("TRUNCATE TABLE " + tableName + ";");
    }

    protected Statement createStatement() throws SQLException {
        return embeddedPostgres.getPostgresDatabase().getConnection().createStatement();
    }

    protected void migrateDatabase(DataSource dataSource) {
        Flyway flyway = Flyway.configure().dataSource(dataSource).load();
        flyway.migrate();
    }
}

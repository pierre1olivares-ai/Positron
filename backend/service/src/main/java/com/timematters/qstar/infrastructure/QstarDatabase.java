package com.timematters.qstar.infrastructure;

import com.timematters.qstar.api.CustomOffsetDateTimeMapper;
import com.timematters.qstar.infrastructure.tools.DatabaseConnection;
import java.time.temporal.ChronoUnit;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.jdbi.v3.core.Jdbi;
import org.jdbi.v3.core.statement.SqlLogger;
import org.jdbi.v3.core.statement.StatementContext;
import org.jdbi.v3.jackson2.Jackson2Config;
import org.jdbi.v3.jackson2.Jackson2Plugin;
import org.jdbi.v3.postgres.PostgresPlugin;
import org.jdbi.v3.sqlobject.SqlObjectPlugin;
import org.springframework.beans.factory.annotation.Value;

/** Unwired IT-template example; the running Q-Star backend does not use a database. */
public class QstarDatabase implements DatabaseConnection {
    private String url;
    private String user;
    private String pass;
    private static final Logger logger = LogManager.getLogger();

    @Value("${app.enablesqllogging}")
    private Boolean enableSQLLogging;

    public QstarDatabase(
            @Value("${spring.datasource.url}") String url,
            @Value("${spring.datasource.username}") String user,
            @Value("${spring.datasource.password}") String pass) {
        this.url = url;
        this.user = user;
        this.pass = pass;
    }

    @Override
    public Jdbi getSqlConnection() {
        Jdbi currentInstance =
                Jdbi.create(url, user, pass)
                        .installPlugin(new PostgresPlugin())
                        .installPlugin(new SqlObjectPlugin())
                        .installPlugin(new Jackson2Plugin());

        currentInstance.getConfig(Jackson2Config.class).setMapper(new CustomOffsetDateTimeMapper());

        if (enableSQLLogging) {
            currentInstance.setSqlLogger(getSqlLogger());
        }

        return currentInstance;
    }

    private SqlLogger getSqlLogger() {
        SqlLogger sqlLogger =
                new SqlLogger() {
                    @Override
                    public void logAfterExecution(StatementContext context) {
                        logger.info(
                                "sql {}, parameters {}, timeTaken {} ms",
                                context.getRenderedSql(),
                                context.getBinding().toString(),
                                context.getElapsedTime(ChronoUnit.MILLIS));
                    }
                };

        return sqlLogger;
    }
}

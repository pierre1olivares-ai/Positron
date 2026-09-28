package com.timematters.qstar.infrastructure.tools;

import org.jdbi.v3.core.Jdbi;

public interface DatabaseConnection {
    Jdbi getSqlConnection();
}

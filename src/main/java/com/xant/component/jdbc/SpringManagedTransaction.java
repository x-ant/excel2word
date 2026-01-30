package com.xant.component.jdbc;

import lombok.extern.slf4j.Slf4j;
import org.apache.ibatis.transaction.Transaction;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.SQLException;

import static cn.hutool.core.lang.Assert.notNull;

@Slf4j
public class SpringManagedTransaction implements Transaction {

    private final DataSource dataSource;

    private Connection connection;

    private boolean isConnectionTransactional;

    private boolean autoCommit;

    public SpringManagedTransaction(DataSource dataSource) {
        notNull(dataSource, "No DataSource specified");
        this.dataSource = dataSource;
    }

    @Override
    public Connection getConnection() throws SQLException {
        if (this.connection == null) {
            openConnection();
        }
        return this.connection;
    }

    private void openConnection() throws SQLException {
        this.connection = ConnectionContext.getConnection(this.dataSource);
        this.autoCommit = this.connection.getAutoCommit();
        this.isConnectionTransactional = TransactionContextManager.getInTransaction();

        log.debug("JDBC Connection [" + this.connection + "] will"
                + (this.isConnectionTransactional ? " " : " not ") + "be managed by Spring");
    }

    @Override
    public void commit() throws SQLException {
        if (isNotInTransaction()) {
            log.debug("Committing JDBC Connection [" + this.connection + "]");
            this.connection.commit();
        }
    }

    @Override
    public void rollback() throws SQLException {
        if (isNotInTransaction()) {
            log.debug("Rolling back JDBC Connection [" + this.connection + "]");
            this.connection.rollback();
        }
    }

    @Override
    public void close() throws SQLException {
        if (isNotInTransaction()) {
            log.debug("Closing JDBC Connection [" + this.connection + "]");
            this.connection.close();
        }
    }

    @Override
    public Integer getTimeout() throws SQLException {
        return null;
    }

    private boolean isNotInTransaction() {
        return this.connection != null && !this.isConnectionTransactional && !this.autoCommit;
    }

}
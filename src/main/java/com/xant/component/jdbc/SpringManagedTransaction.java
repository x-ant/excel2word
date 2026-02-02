package com.xant.component.jdbc;

import lombok.extern.slf4j.Slf4j;
import org.apache.ibatis.transaction.Transaction;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.SQLException;

import static cn.hutool.core.lang.Assert.notNull;

/**
 * 代理的实际是数据库连接
 * 功能
 * 1、从线程变量中获取数据库连接，如果有则复用没有则创建
 * 2、数据库连接的操作补充事务判断，如果在事务中则不进行操作，由事务管理器进行控制
 *
 * @author xuhq
 */
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

    /**
     * 当前数据库连接不在事务中
     * 1、数据库连接不为空
     * 2、数据库连接不是自动提交，否则就不用操作了
     * 3、数据库连接不在事务中，事务中的由事务控制
     *
     * @return boolean
     */
    private boolean isNotInTransaction() {
        return this.connection != null && !this.isConnectionTransactional && !this.autoCommit;
    }

}
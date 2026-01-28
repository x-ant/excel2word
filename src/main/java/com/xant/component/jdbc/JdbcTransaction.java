package com.xant.component.jdbc;

import org.apache.ibatis.transaction.Transaction;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.SQLException;

/**
 * 事务管理器
 *
 * @author xuhq
 */
public class JdbcTransaction implements Transaction {

    private final DataSource dataSource;

    /**
     * 必须传入一个数据库对象
     *
     * @param dataSource 数据库实例
     */
    public JdbcTransaction(DataSource dataSource) {
        this.dataSource = dataSource;
    }

    public Connection getConnection() throws SQLException {
        return ConnectionContext.getConnection(dataSource);
    }

    public void commit() throws SQLException {
        Connection connection = getConnection();
        if (!connection.getAutoCommit()) {
            connection.commit();
        }
    }

    public void rollback() throws SQLException {
        Connection connection = getConnection();
        if (!connection.getAutoCommit()) {
            connection.rollback();
        }
    }

    public void close() throws SQLException {
        Connection connection = getConnection();
        if (!connection.getAutoCommit()) {
            connection.setAutoCommit(true);
        }
        connection.close();
    }

    @Override
    public Integer getTimeout() throws SQLException {
        return null;
    }

}

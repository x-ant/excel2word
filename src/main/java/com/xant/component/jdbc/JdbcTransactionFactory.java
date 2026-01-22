package com.xant.component.jdbc;

import org.apache.ibatis.session.TransactionIsolationLevel;
import org.apache.ibatis.transaction.Transaction;
import org.apache.ibatis.transaction.TransactionFactory;

import javax.sql.DataSource;
import java.sql.Connection;
import java.util.Properties;

/**
 * JDBC事务工厂
 *
 * @author xuhq
 */
public class JdbcTransactionFactory implements TransactionFactory {

    @Override
    public Transaction newTransaction(DataSource dataSource, TransactionIsolationLevel level, boolean autoCommit) {
        return new JdbcTransaction(dataSource);
    }

    @Override
    public Transaction newTransaction(Connection conn) {
        throw new UnsupportedOperationException("New Jdbc transactions require a DataSource");
    }

    @Override
    public void setProperties(Properties props) {
        // not needed in this version
    }

}
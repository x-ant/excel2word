package com.xant.component.jdbc;

import lombok.extern.slf4j.Slf4j;

import java.sql.Connection;
import java.util.List;
import java.util.Objects;
import java.util.function.Supplier;

/**
 * 事务工具类
 *
 * @author xuhq
 */
@Slf4j
public class TransactionUtil {

    private TransactionUtil() {
    }

    public static <T> T transactionWithRequired(Supplier<T> supplier) {
        T result = null;
        Connection connection = null;
        Boolean originAutoCommit = null;
        try {
            connection = ConnectionContext.getConnection(SqliteSingleton.getSingleton());
            originAutoCommit = connection.getAutoCommit();
            if (originAutoCommit) {
                connection.setAutoCommit(false);
            }
            TransactionContextManager.setInTransaction(true);

            result = supplier.get();

            commitCallback();
            connection.commit();
        } catch (Throwable tr) {
            try {
                if (Objects.nonNull(originAutoCommit)) {
                    rollbackCallback();
                    connection.rollback();
                }
            } catch (Throwable tr2) {
                log.error("数据库事务回滚失败！", tr2);
            }

            throw new RuntimeException(tr);
        } finally {
            try {
                ConnectionContext.removeConnection(SqliteSingleton.getSingleton());
                if (Objects.nonNull(originAutoCommit)) {
                    connection.setAutoCommit(originAutoCommit);
                    connection.close();
                }
                TransactionContextManager.setInTransaction(false);
            } catch (Throwable tr2) {
                log.error("数据库连接关闭失败！", tr2);
            }
        }
        return result;
    }

    private static void commitCallback() {
        List<TransactionCallback> transactionCallback = TransactionContextManager.getTransactionCallback();
        for (TransactionCallback callback : transactionCallback) {
            callback.beforeCommit();
        }
        for (TransactionCallback callback : transactionCallback) {
            callback.beforeCompletion();
        }
    }

    private static void rollbackCallback() {
        List<TransactionCallback> transactionCallback = TransactionContextManager.getTransactionCallback();
        for (TransactionCallback callback : transactionCallback) {
            callback.beforeCompletion();
        }
    }

}

package com.xant.component.jdbc;

import org.apache.ibatis.session.SqlSession;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class TransactionContextManager {

    private static final String IN_TRANSACTION_KEY = "IN_TRANSACTION_KEY";

    private static final String SQL_SESSION_KEY = "SQL_SESSION_KEY";

    private static final String TRANSACTION_CALLBACK_KEY = "TRANSACTION_CALLBACK_KEY";

    private TransactionContextManager() {
    }

    public static void setInTransaction(boolean inTransaction) {
        ThreadContext.setAttribute(IN_TRANSACTION_KEY, inTransaction);
    }

    public static boolean getInTransaction() {
        return Boolean.TRUE.equals(ThreadContext.getAttribute(IN_TRANSACTION_KEY));
    }

    public static void clearInTransaction() {
        ThreadContext.removeAttribute(IN_TRANSACTION_KEY);
    }

    public static SqlSession getSqlSession() {
        return ThreadContext.getAttribute(SQL_SESSION_KEY);
    }

    public static void setSqlSession(SqlSession sqlSession) {
        ThreadContext.setAttribute(SQL_SESSION_KEY, sqlSession);
    }

    public static SqlSession clearSqlSession() {
        return ThreadContext.removeAttribute(SQL_SESSION_KEY);
    }

    public static List<TransactionCallback> getTransactionCallback() {
        List<TransactionCallback> callbackList = ThreadContext.getAttribute(TRANSACTION_CALLBACK_KEY);
        if (Objects.isNull(callbackList)) {
            callbackList = new ArrayList<>();
            ThreadContext.setAttribute(TRANSACTION_CALLBACK_KEY, callbackList);
        }
        return callbackList;
    }

    public static void setTransactionCallback(List<TransactionCallback> callbackList) {
        getTransactionCallback().addAll(callbackList);
    }

    public static List<TransactionCallback> clearTransactionCallback() {
        List<TransactionCallback> callbackList = getTransactionCallback();
        callbackList.clear();
        return callbackList;
    }

}

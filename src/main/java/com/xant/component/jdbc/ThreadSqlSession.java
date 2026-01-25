package com.xant.component.jdbc;

import com.xant.component.SqlSessionFactorySingleton;
import org.apache.ibatis.session.ExecutorType;
import org.apache.ibatis.session.SqlSession;

import java.util.Objects;

public class ThreadSqlSession {

    private static ThreadLocal<SqlSession> sqlSessionThreadLocal = new ThreadLocal<>();
    private static ThreadLocal<Boolean> isTransactionThreadLocal = new ThreadLocal<>();

    private static final ExecutorType executorType = ExecutorType.SIMPLE;

    private ThreadSqlSession() {
    }

    public static SqlSession getOrCreateSqlSession() {
        SqlSession sqlSession = sqlSessionThreadLocal.get();
        if (Objects.isNull(sqlSession)) {
            sqlSession = SqlSessionFactorySingleton.getSingleton().openSession(ThreadSqlSession.executorType);
            sqlSessionThreadLocal.set(sqlSession);
        }
        return sqlSession;
    }

    public static void clearSqlSession() {
        sqlSessionThreadLocal.remove();
    }

    public static boolean isInTransaction() {
        return Boolean.TRUE.equals(isTransactionThreadLocal.get());
    }

}

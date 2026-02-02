package com.xant.component.jdbc;

import lombok.extern.slf4j.Slf4j;
import org.apache.ibatis.session.SqlSession;

/**
 * 事务回调方法，时机包括 开始、提交、回滚、关闭
 *
 * @author xuhq
 */
@Slf4j
public class SqlSessionTransactionCallback implements TransactionCallback {

    private SqlSession sqlSession;

    public SqlSessionTransactionCallback(SqlSession sqlSession) {
        this.sqlSession = sqlSession;
    }

    @Override
    public void beforeCommit() {
        if (TransactionContextManager.getInTransaction()) {
            sqlSession.commit();
        }
    }

    @Override
    public void beforeCompletion() {
        if (TransactionContextManager.getInTransaction()) {
            TransactionContextManager.clearSqlSession();
            sqlSession.close();
        }
    }
}

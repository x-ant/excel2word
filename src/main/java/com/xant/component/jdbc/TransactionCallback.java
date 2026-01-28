package com.xant.component.jdbc;

/**
 * 事务操作回调
 *
 * @author xuhq
 */
public interface TransactionCallback {

    void beforeCommit();

    void beforeCompletion();
}

package com.xant.util;

import com.xant.component.SqlSessionManagerSingleton;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.ibatis.session.SqlSessionManager;

import java.util.Objects;
import java.util.function.Supplier;

/**
 * 编程式事务管理
 *
 * @author xuhq
 */
@Slf4j
public class TransactionUtil {

    private TransactionUtil() {
    }

    @SneakyThrows
    public static <T> T doTransactionWithRequired(Supplier<T> supplier) {
        SqlSessionManager transactionManager = SqlSessionManagerSingleton.getSingleton();
        transactionManager.startManagedSession();

        T result = null;
        Throwable executeException = null;
        try {
            result = supplier.get();
        } catch (Throwable tr1) {
            executeException = tr1;
        }

        if (!Objects.isNull(executeException)) {
            try {
                transactionManager.rollback();
            } catch (Throwable tr2) {
                log.error("数据库事务回滚失败！", tr2);
            }
            throw executeException;
        } else {
            transactionManager.commit();
        }
        return result;
    }

}
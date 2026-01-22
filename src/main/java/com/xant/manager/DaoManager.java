package com.xant.manager;

import lombok.extern.slf4j.Slf4j;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 年度账单管理
 *
 * @author xuhq
 */
@Slf4j
public class DaoManager {

    private Map<Class<?>, Object> daoProxyMap = new ConcurrentHashMap<>();

    private DaoManager() {
    }

    private static class DaoManagerHolder {
        private static final DaoManager INSTANCE = new DaoManager();
    }

    public static DaoManager getInstance() {
        return DaoManagerHolder.INSTANCE;
    }

    public <T> T getDaoManager(Class<T> clazz) {

        /*InvocationHandler invocationHandler = (proxy, method, args) -> {
            if (Object.class.equals(method.getDeclaringClass())) {
                return method.invoke(this, args);
            }
            return daoProxyMap(method).invoke(proxy, method, args, sqlSession);
        };
        Object proxyInstance = Proxy.newProxyInstance(clazz.getClassLoader(), clazz.getInterfaces(), invocationHandler);

        SqlSessionFactory sqlSessionFactory = SqlSessionFactorySingleton.getSingleton();
        SqlSession sqlSession = sqlSessionFactory.openSession();
        try {
            YearBillMapper yearBillMapper = sqlSession.getMapper(YearBillMapper.class);
            yearBillMapper.insert(yearBillPO);
            sqlSession.commit();
        } catch (Exception e) {
            sqlSession.rollback();
            log.error("新增异常", e);
        } finally {
            sqlSession.close();
        }*/
        return null;
    }


}

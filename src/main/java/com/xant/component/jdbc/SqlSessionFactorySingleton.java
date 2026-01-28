package com.xant.component.jdbc;

import lombok.Data;
import org.apache.ibatis.io.Resources;
import org.apache.ibatis.session.ExecutorType;
import org.apache.ibatis.session.SqlSession;
import org.apache.ibatis.session.SqlSessionFactory;
import org.apache.ibatis.session.SqlSessionFactoryBuilder;

import java.io.InputStream;

/**
 * SqlSessionFactory单例
 *
 * @author xuhq
 */
@Data
public class SqlSessionFactorySingleton {

    private static final String CONFIG_FILE = "mybatis-config.xml";

    private static volatile SqlSessionFactory instance;

    private static final ExecutorType executorType = ExecutorType.SIMPLE;

    private SqlSessionFactorySingleton() {
    }

    public static SqlSessionFactory getSingleton() {
        if (instance == null) {
            synchronized (SqlSessionFactorySingleton.class) {
                if (instance == null) {
                    try (InputStream inputStream = Resources.getResourceAsStream(CONFIG_FILE)) {
                        instance = new SqlSessionFactoryBuilder().build(inputStream);
                    } catch (Exception e) {
                        throw new RuntimeException("数据库配置文件读取失败，请联系管理员！");
                    }
                }
            }
        }
        return instance;
    }

    public static SqlSession getSqlSession(ExecutorType executorType) {
        return getSingleton().openSession(executorType);
    }

    public static SqlSession getSqlSession() {
        return getSingleton().openSession(executorType);
    }

}

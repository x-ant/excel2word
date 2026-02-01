package com.xant.component.jdbc;

import lombok.Data;
import org.apache.ibatis.io.Resources;
import org.apache.ibatis.session.ExecutorType;
import org.apache.ibatis.session.SqlSession;
import org.apache.ibatis.session.SqlSessionFactory;
import org.apache.ibatis.session.SqlSessionFactoryBuilder;

import java.io.InputStream;
import java.util.Objects;

/**
 * SqlSessionFactory单例
 *
 * @author xuhq
 */
@Data
public class SqlSessionFactorySingleton {

    private static final String CONFIG_FILE = "mybatis-config.xml";

    private static volatile SqlSessionFactory sqlSessionFactory;
    private static volatile SqlSessionTemplate batchSqlSession;
    private static volatile SqlSessionTemplate singleSqlSession;

    private SqlSessionFactorySingleton() {
    }

    public static SqlSessionFactory getSqlSessionFactory() {
        if (Objects.isNull(sqlSessionFactory)) {
            synchronized (SqlSessionFactorySingleton.class) {
                if (Objects.isNull(sqlSessionFactory)) {
                    try (InputStream inputStream = Resources.getResourceAsStream(CONFIG_FILE)) {
                        sqlSessionFactory = new SqlSessionFactoryBuilder().build(inputStream);
                    } catch (Exception e) {
                        throw new RuntimeException("数据库配置文件读取失败，请联系管理员！", e);
                    }
                }
            }
        }
        return sqlSessionFactory;
    }

    public static SqlSession getBatchSqlSession() {
        if (Objects.isNull(batchSqlSession)) {
            synchronized (SqlSessionFactorySingleton.class) {
                if (Objects.isNull(batchSqlSession)) {
                    batchSqlSession = new SqlSessionTemplate(ExecutorType.BATCH);
                }
            }
        }
        return batchSqlSession;
    }

    public static SqlSession getSingleSqlSession() {
        if (Objects.isNull(singleSqlSession)) {
            synchronized (SqlSessionFactorySingleton.class) {
                if (Objects.isNull(singleSqlSession)) {
                    singleSqlSession = new SqlSessionTemplate(ExecutorType.SIMPLE);
                }
            }
        }
        return singleSqlSession;
    }

}

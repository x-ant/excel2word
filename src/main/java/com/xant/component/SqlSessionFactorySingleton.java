package com.xant.component;

import lombok.Data;
import org.apache.ibatis.io.Resources;
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

    private static final String CONFIG_FILE = "chapter1/mybatis-config.xml";

    private static volatile SqlSessionFactory instance;

    private SqlSessionFactorySingleton() {
    }

    public static SqlSessionFactory getSingleton() {
        if (instance == null) {
            synchronized (SqlSessionFactorySingleton.class) {
                if (instance == null) {
                    try (InputStream inputStream = Resources.getResourceAsStream(CONFIG_FILE)) {
                        instance = new SqlSessionFactoryBuilder().build(inputStream);
                    } catch (Exception e) {
                        throw new RuntimeException("数据工厂初始化失败，请联系管理员！");
                    }
                }
            }
        }
        return instance;
    }

}

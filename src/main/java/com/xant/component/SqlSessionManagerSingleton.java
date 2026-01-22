package com.xant.component;

import lombok.Data;
import org.apache.ibatis.io.Resources;
import org.apache.ibatis.session.SqlSession;
import org.apache.ibatis.session.SqlSessionManager;

import java.io.InputStream;

/**
 * SqlSessionManager单例
 *
 * @author xuhq
 */
@Data
public class SqlSessionManagerSingleton {

    private static final String CONFIG_FILE = "chapter1/mybatis-config.xml";

    private static volatile SqlSessionManager instance;

    private SqlSessionManagerSingleton() {
    }

    public static SqlSessionManager getSingleton() {
        if (instance == null) {
            synchronized (SqlSessionManagerSingleton.class) {
                if (instance == null) {
                    try (InputStream inputStream = Resources.getResourceAsStream(CONFIG_FILE)) {
                        instance = SqlSessionManager.newInstance(inputStream);
                    } catch (Exception e) {
                        throw new RuntimeException("数据库配置文件读取失败，请联系管理员！");
                    }
                }
            }
        }
        return instance;
    }

}

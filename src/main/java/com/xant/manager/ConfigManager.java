package com.xant.manager;

import cn.hutool.core.bean.BeanUtil;
import com.xant.entity.ConfigPO;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;

import java.io.*;
import java.util.Map;
import java.util.Properties;

/**
 * 配置实体类管理
 *
 * @author xuhq
 */
@Slf4j
public class ConfigManager {

    /**
     * 项目根目录
     * 程序启动的路径
     */
    private static final String PROJECT_ROOT = System.getProperty("user.dir");

    private static class SingletonHolder {
        private static ConfigPO INSTANCE = new ConfigPO();

        static {
            Properties properties = new Properties();
            File configFile = new File(PROJECT_ROOT, "config.properties");
            if (configFile.exists()) {
                try {
                    properties.load(new FileReader(configFile));
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
            BeanUtil.fillBeanWithMap(properties, INSTANCE, true);
        }
    }

    private ConfigManager() {
    }

    public static ConfigPO getSingletonConfigPO() {
        return SingletonHolder.INSTANCE;
    }

    public static ConfigPO getPrototypeConfigPO() {
        return BeanUtil.copyProperties(SingletonHolder.INSTANCE, ConfigPO.class);
    }

    @SneakyThrows
    public static void setConfigPO(ConfigPO configPO) {
        Map<String, Object> field2ValueMap = BeanUtil.beanToMap(configPO);
        Properties properties = new Properties();
        properties.putAll(field2ValueMap);
        try (OutputStream output = new FileOutputStream("config.properties")) {
            properties.store(output, "Config文件");
        }
        SingletonHolder.INSTANCE = configPO;
    }

}

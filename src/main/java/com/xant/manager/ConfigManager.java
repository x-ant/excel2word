package com.xant.manager;

import cn.hutool.core.bean.BeanUtil;
import cn.hutool.core.io.FileUtil;
import cn.hutool.json.JSONUtil;
import com.xant.entity.ConfigPO;
import lombok.extern.slf4j.Slf4j;

import java.io.File;
import java.nio.charset.StandardCharsets;

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
    private static final String CONFIG_FILE = PROJECT_ROOT + File.separator + "config.json";

    private static class SingletonHolder {
        private static ConfigPO INSTANCE = new ConfigPO();

        static {
            File configFile = new File(CONFIG_FILE);
            log.info("读取配置文件路径: {}", configFile.getAbsolutePath());
            if (configFile.exists()) {
                String configJsonStr = FileUtil.readString(configFile, StandardCharsets.UTF_8);
                BeanUtil.fillBeanWithMap(JSONUtil.parseObj(configJsonStr), INSTANCE, true);
            }
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

    public static void setConfigPO(ConfigPO configPO) {
        File configFile = new File(CONFIG_FILE);
        if (!FileUtil.exist(configFile)) {
            FileUtil.touch(configFile);
        }
        log.info("写入配置文件路径: {}", configFile.getAbsolutePath());
        FileUtil.writeString(JSONUtil.toJsonPrettyStr(configPO), configFile, StandardCharsets.UTF_8);
        BeanUtil.copyProperties(configPO, SingletonHolder.INSTANCE, true);
    }

}

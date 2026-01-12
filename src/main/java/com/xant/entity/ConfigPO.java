package com.xant.entity;

import lombok.Data;

/**
 * 配置信息文件
 *
 * @author xuhq
 */
@Data
public class ConfigPO {

    private String templatePrefix = "${";

    private String templateSuffix = "}";

    private String templateFile;

    private String inputDir;

    private String outputDir;

}

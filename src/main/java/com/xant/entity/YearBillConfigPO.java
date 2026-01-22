package com.xant.entity;

import lombok.Data;

/**
 * 年度账单配置信息
 *
 * @author xant
 */
@Data
public class YearBillConfigPO {

    /**
     * 输入文件
     */
    private String inputFile;

    /**
     * 是否按名称排序
     */
    private Boolean inputFileIsOrderByName = true;

    /**
     * 输入文件年份列
     */
    private Integer inputFileYearColIndex;

    /**
     * 输入文件对象名称列
     */
    private Integer inputFileNameColIndex;

    /**
     * 输入文件年度收入金额列
     */
    private Integer inputFileAmountColIndex;

    /**
     * 输入文件年度结余金额列
     */
    private Integer inputFileBalanceColIndex;

    /**
     * 输出文件
     */
    private String outputFile;

    /**
     * 输出文件sheet名称
     */
    private String outputFileSheetName;

    /**
     * 输出文件账龄填充开始列
     */
    private Integer outputFileBillAgeFillStartColIndex;

}

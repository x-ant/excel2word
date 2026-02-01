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
    private String inputFile = "D:\\exe4j\\辅助核算汇总表查询结果.xls";

    /**
     * 是否按名称排序
     */
    private Boolean inputFileIsOrderByName = false;

    /**
     * 输入文件年份列
     */
    private String inputFileYearCol = "A";

    /**
     * 输入文件对象名称列
     */
    private String inputFileNameCol = "G";

    /**
     * 输入文件年度收入金额列
     */
    private String inputFileAmountCol = "J";

    /**
     * 输入文件年度结余金额列
     */
    private String inputFileBalanceCol = "M";

    /**
     * 输出文件
     */
    private String outputFile = "D:\\exe4j\\4280其他应收款基础表格(1).xls";

    /**
     * 输出文件sheet名称
     */
    private String outputFileSheetName = "明细表";

    /**
     * 输出文件账龄填充开始列
     */
    private String outputFileBillAgeFillStartCol = "P";

}

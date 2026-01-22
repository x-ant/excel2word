package com.xant.entity;

import lombok.Data;

import java.math.BigDecimal;

/**
 * 年度账单信息
 *
 * @author xuhq
 */
@Data
public class YearBillPO {

    /**
     * 主键
     */
    private String id;

    /**
     * 年份
     */
    private Integer year;

    /**
     * 对象名称
     */
    private String company;

    /**
     * 年度收入金额
     */
    private BigDecimal recAmount;

    /**
     * 年度结余金额
     */
    private BigDecimal balance;

}

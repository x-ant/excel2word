package com.xant.dao;

import com.xant.entity.YearBillPO;
import org.apache.ibatis.annotations.Param;

import java.util.List;

/**
 * 年度账单信息
 *
 * @author xuhq
 */
public interface YearBillMapper extends BaseMapper<YearBillPO> {

    YearBillPO selectById(String id);

    int insert(YearBillPO yearBillPO);

    int updateById(YearBillPO yearBillPO);

    int deleteById(String id);

    int deleteByIdList(@Param("idList") List<String> idList);

}

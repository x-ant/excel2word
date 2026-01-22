package com.xant.dao;

import org.apache.ibatis.annotations.Param;

import java.util.List;

/**
 * 基类Mapper
 *
 * @author xuhq
 */
public interface BaseMapper<T> {

    T selectById(String id);

    int insert(T po);

    int updateById(T po);

    int deleteById(String id);

    int deleteByIdList(@Param("idList") List<String> idList);

}

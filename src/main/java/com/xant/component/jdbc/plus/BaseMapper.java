package com.xant.component.jdbc.plus;

import com.xant.common.constant.BaseSqlConstant;
import org.apache.ibatis.annotations.*;

import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * 基类Mapper
 *
 * @author xuhq
 */
public interface BaseMapper<T> {

    @SelectProvider(type = BaseSqlProvider.class, method = "selectById")
    T selectById(@Param("id") String id);

    @InsertProvider(type = BaseSqlProvider.class, method = "insert")
    int insert(@Param(BaseSqlConstant.ENTITY) T entity);

    @UpdateProvider(type = BaseSqlProvider.class, method = "updateById")
    int updateById(@Param(BaseSqlConstant.ENTITY) T entity);

    @DeleteProvider(type = BaseSqlProvider.class, method = "deleteById")
    int deleteById(@Param(BaseSqlConstant.ID) String id);

    /**
     * 按ID批量查询
     *
     * @param idList id列表
     * @return 结果集
     */
    @SelectProvider(type = BaseSqlProvider.class, method = "selectByIdList")
    List<T> selectByIdList(@Param("idList") Collection<String> idList);

    /**
     * 按条件删除
     *
     * @param fieldMap 查询条件
     * @return 结果集
     */
    @DeleteProvider(type = BaseSqlProvider.class, method = "deleteByMap")
    int deleteByMap(@Param("fieldMap") Map<String, Object> fieldMap);

    /**
     * 按条件查询
     *
     * @param fieldMap 查询条件
     * @return 结果集
     */
    @SelectProvider(type = BaseSqlProvider.class, method = "selectByMap")
    List<T> selectByMap(@Param(BaseSqlConstant.FIELD_MAP) Map<String, Object> fieldMap,
                        @Param(BaseSqlConstant.ORDER_BY_LIST) List<String> orderByList,
                        @Param(BaseSqlConstant.OFFSET) Integer offset,
                        @Param(BaseSqlConstant.LIMIT) Integer limit);

    /**
     * 按条件查询
     *
     * @param entity 查询条件
     * @return 结果集
     */
    @SelectProvider(type = BaseSqlProvider.class, method = "selectByEntity")
    List<T> selectByEntity(@Param(BaseSqlConstant.ENTITY) T entity);

    /**
     * 批量删除
     *
     * @param idList id列表
     * @return 受影响行数
     */
    @DeleteProvider(type = BaseSqlProvider.class, method = "deleteByIdList")
    int deleteByIdList(@Param("idList") Collection<String> idList);


    /**
     * 分页查询
     *
     * @param entity  查询条件
     * @param offset  偏移量
     * @param limit   查询数量
     * @param orderBy 排序字段
     * @return 结果集
     */
    @SelectProvider(type = BaseSqlProvider.class, method = "selectPage")
    List<T> selectPage(@Param("entity") T entity,
                       @Param("offset") Integer offset,
                       @Param("limit") Integer limit,
                       @Param("orderBy") List<String> orderBy);

}

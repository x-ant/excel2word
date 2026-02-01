package com.xant.component.jdbc.plus;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import com.xant.component.jdbc.TransactionUtil;
import com.xant.dao.BaseMapper;

import java.util.Collection;
import java.util.List;
import java.util.Map;

public interface IService<T> {

    /**
     * 默认批次提交数量
     */
    int DEFAULT_BATCH_SIZE = 1000;

    /**
     * 插入一条记录（选择字段，策略插入）
     *
     * @param entity 实体对象
     */
    default boolean save(T entity) {
        return retBool(getBaseMapper().insert(entity));
    }

    /**
     * 插入（批量）
     *
     * @param entityList 实体对象集合
     */
    default boolean saveBatch(Collection<T> entityList) {
        return TransactionUtil.transactionWithRequired(() -> {
            return saveBatch(entityList, DEFAULT_BATCH_SIZE);
        });
    }

    /**
     * 插入（批量）
     *
     * @param entityList 实体对象集合
     * @param batchSize  插入批次数量
     */
    boolean saveBatch(Collection<T> entityList, int batchSize);

    /**
     * TableId 注解存在更新记录，否插入一条记录
     *
     * @param entity 实体对象
     */
    boolean saveOrUpdate(T entity);

    /**
     * 批量修改插入
     *
     * @param entityList 实体对象集合
     */
    default boolean saveOrUpdateBatch(Collection<T> entityList) {
        return TransactionUtil.transactionWithRequired(() -> {
            return saveOrUpdateBatch(entityList, DEFAULT_BATCH_SIZE);
        });
    }

    /**
     * 批量修改插入
     *
     * @param entityList 实体对象集合
     * @param batchSize  每次的数量
     */
    boolean saveOrUpdateBatch(Collection<T> entityList, int batchSize);

    /**
     * 根据 ID 删除
     *
     * @param id 主键ID
     */
    default boolean removeById(String id) {
        return retBool(getBaseMapper().deleteById(id));
    }

    /**
     * 删除（根据ID 批量删除）
     *
     * @param idList 主键ID列表
     */
    default boolean removeByIdList(Collection<String> idList) {
        if (CollUtil.isEmpty(idList)) {
            return false;
        }
        return retBool(getBaseMapper().deleteByIdList(idList));
    }

    /**
     * 根据 columnMap 条件，删除记录
     *
     * @param columnMap 表字段 map 对象
     */
    default boolean removeByMap(Map<String, Object> columnMap) {
        Assert.notEmpty(columnMap, "error: columnMap must not be empty");
        return retBool(getBaseMapper().deleteByMap(columnMap));
    }

    /**
     * 根据 ID 选择修改
     *
     * @param entity 实体对象
     */
    default boolean updateById(T entity) {
        return retBool(getBaseMapper().updateById(entity));
    }

    /**
     * 根据ID 批量更新
     *
     * @param entityList 实体对象集合
     */
    default boolean updateBatchById(Collection<T> entityList) {
        return TransactionUtil.transactionWithRequired(() -> {
            return updateBatchById(entityList, DEFAULT_BATCH_SIZE);
        });
    }

    /**
     * 根据ID 批量更新
     *
     * @param entityList 实体对象集合
     * @param batchSize  更新批次数量
     */
    boolean updateBatchById(Collection<T> entityList, int batchSize);

    /**
     * 根据 ID 查询
     *
     * @param id 主键ID
     */
    default T selectById(String id) {
        return getBaseMapper().selectById(id);
    }

    /**
     * 查询（根据ID 批量查询）
     *
     * @param idList 主键ID列表
     */
    default List<T> selectByIdList(Collection<String> idList) {
        return getBaseMapper().selectByIdList(idList);
    }

    /**
     * 查询（根据 columnMap 条件）
     *
     * @param columnMap 表字段 map 对象
     */
    default List<T> selectByMap(Map<String, Object> columnMap) {
        return getBaseMapper().selectByMap(columnMap);
    }

    /**
     * 获取对应 entity 的 BaseMapper
     *
     * @return BaseMapper
     */
    BaseMapper<T> getBaseMapper();

    /**
     * 判断数据库操作是否成功
     *
     * @param result 数据库操作返回影响条数
     * @return boolean
     */
    default boolean retBool(Integer result) {
        return null != result && result >= 1;
    }

    /**
     * 返回SelectCount执行结果
     *
     * @param result ignore
     * @return int
     */
    default int retCount(Integer result) {
        return (null == result) ? 0 : result;
    }

}
package com.xant.component.jdbc.plus;

import cn.hutool.core.collection.CollUtil;
import com.xant.util.UUIDUtil;
import org.apache.ibatis.annotations.Param;
import org.apache.ibatis.jdbc.SQL;

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * 通用SQL提供器
 *
 * @author xuhq
 */
public class BaseSqlProvider {

    /**
     * 使用元数据缓存
     *
     * @param clazz 实体类
     * @return 元数据
     */
    private EntityMetaCache.EntityMeta getEntityMeta(Class<?> clazz) {
        return EntityMetaCache.getEntityMeta(clazz);
    }

    /**
     * 通用的 selectById 方法
     *
     * @param entity 实体对象
     * @return SQL
     */
    public String selectById(Object entity) {
        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        if (meta.getIdColumn() == null) {
            throw new RuntimeException("实体类没有定义主键字段");
        }

        SQL sql = new SQL()
                .SELECT(getSelectColumns(meta))
                .FROM(meta.getTableName())
                .WHERE(meta.getIdColumn().getColumnName() + " = #{id}");

        return sql.toString();
    }

    /**
     * 通用的 selectById 方法
     *
     * @param idList 主键ID列表
     * @param entity 实体对象
     * @return SQL
     */
    public String selectByIdList(@Param("idList") Collection<String> idList, Object entity) {
        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        if (meta.getIdColumn() == null) {
            throw new RuntimeException("实体类没有定义主键字段");
        }

        SQL sql = new SQL()
                .SELECT(getSelectColumns(meta))
                .FROM(meta.getTableName())
                .WHERE(meta.getIdColumn().getColumnName() + " IN (#{idList})");

        return sql.toString();
    }

    /**
     * 通用的INSERT方法
     *
     * @param entity 实体对象
     * @return SQL
     */
    public String insert(Object entity) {
        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        SQL sql = new SQL().INSERT_INTO(meta.getTableName());

        for (EntityMetaCache.ColumnMeta column : meta.getColumnList()) {
            // 跳过非插入字段
            if (!column.isInsertable()) {
                continue;
            }

            // 主键
            if (column.isId()) {
                column.setFieldValue(entity, UUIDUtil.getUUID());
            }

            sql.VALUES(column.getColumnName(), "#{" + column.getFieldName() + "}");
        }

        return sql.toString();
    }

    /**
     * 通用的 updateById 方法
     *
     * @param entity 实体对象
     * @return SQL
     */
    public String updateById(Object entity) {
        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        if (meta.getIdColumn() == null) {
            throw new RuntimeException("实体类没有定义主键字段");
        }

        SQL sql = new SQL().UPDATE(meta.getTableName());

        for (EntityMetaCache.ColumnMeta column : meta.getColumnList()) {
            // 跳过主键和非更新字段
            if (column.isId() || !column.isUpdatable()) {
                continue;
            }

            sql.SET(column.getColumnName() + " = #{" + column.getFieldName() + "}");
        }

        sql.WHERE(meta.getIdColumn().getColumnName() + " = #{id}");

        return sql.toString();
    }

    /**
     * 通用的 deleteById 方法
     *
     * @param entity 实体对象
     * @return SQL
     */
    public String deleteById(@Param("id") String id, Object entity) {
        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        if (meta.getIdColumn() == null) {
            throw new RuntimeException("实体类没有定义主键字段");
        }

        return new SQL()
                .DELETE_FROM(meta.getTableName())
                .WHERE(meta.getIdColumn().getColumnName() + " = #{id}")
                .toString();
    }

    /**
     * 动态条件查询
     *
     * @param params 查询参数
     * @return SQL
     */
    public String selectByMap(Map<String, Object> params, Object entity) {
        Map<String, Object> fieldMap = (Map<String, Object>) params.get("fieldMap");
        if (CollUtil.isEmpty(fieldMap)) {
            throw new RuntimeException("查询条件不能为空");
        }

        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        SQL sql = new SQL()
                .SELECT(getSelectColumns(meta))
                .FROM(meta.getTableName());

        // 动态WHERE条件
        for (Map.Entry<String, Object> field : fieldMap.entrySet()) {
            EntityMetaCache.ColumnMeta columnMeta = meta.getColumn(field.getKey());
            if (Objects.isNull(columnMeta)) {
                continue;
            }
            addWhereCondition(sql, columnMeta, field.getValue());
        }

        // 排序
        addOrderBy(sql, params);

        // 分页
        addPagination(sql, params);

        return sql.toString();
    }

    /**
     * 动态条件查询
     *
     * @param params 查询参数
     * @return SQL
     */
    public String deleteByMap(Map<String, Object> params, Object entity) {
        Map<String, Object> fieldMap = (Map<String, Object>) params.get("fieldMap");
        if (CollUtil.isEmpty(fieldMap)) {
            throw new RuntimeException("查询条件不能为空");
        }

        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        SQL sql = new SQL()
                .DELETE_FROM(meta.getTableName());

        // 动态WHERE条件
        for (Map.Entry<String, Object> field : fieldMap.entrySet()) {
            EntityMetaCache.ColumnMeta columnMeta = meta.getColumn(field.getKey());
            if (Objects.isNull(columnMeta)) {
                continue;
            }
            addWhereCondition(sql, columnMeta, field.getValue());
        }

        return sql.toString();
    }

    /**
     * 动态条件查询
     *
     * @param params 查询参数
     * @return SQL
     */
    public String selectByCondition(Map<String, Object> params) {
        Object entity = params.get("entity");
        if (entity == null) {
            throw new RuntimeException("查询条件不能为空");
        }

        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        SQL sql = new SQL()
                .SELECT(getSelectColumns(meta))
                .FROM(meta.getTableName());

        // 动态WHERE条件
        for (EntityMetaCache.ColumnMeta column : meta.getColumnList()) {
            try {
                Object value = column.getFieldValue(entity);

                if (value != null) {
                    addWhereCondition(sql, column, value);
                }
            } catch (Exception e) {
                throw new RuntimeException("无法访问字段: " + column.getFieldName(), e);
            }
        }

        // 排序
        addOrderBy(sql, params);

        // 分页
        addPagination(sql, params);

        return sql.toString();
    }

    /**
     * 批量插入
     */
    public String deleteByIdList(@Param("idList") Collection<String> idList, Object entity) {
        if (CollUtil.isEmpty(idList)) {
            throw new RuntimeException("批量删除ID不能为空");
        }

        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        SQL sql = new SQL()
                .DELETE_FROM(meta.getTableName())
                .WHERE(meta.getIdColumn().getColumnName() + " IN (#{idList})");

        return sql.toString();
    }

    /**
     * 分页查询
     *
     * @param params 查询参数
     * @return SQL
     */
    public String selectPage(Map<String, Object> params) {
        Object entity = params.get("entity");
        Integer offset = (Integer) params.get("offset");
        Integer limit = (Integer) params.get("limit");

        EntityMetaCache.EntityMeta meta = getEntityMeta(entity.getClass());

        SQL sql = new SQL()
                .SELECT(getSelectColumns(meta))
                .FROM(meta.getTableName());

        // WHERE条件
        for (EntityMetaCache.ColumnMeta column : meta.getColumnList()) {
            try {
                Object value = column.getFieldValue(entity);

                if (value != null) {
                    addWhereCondition(sql, column, value);
                }
            } catch (Exception e) {
                throw new RuntimeException("无法访问字段: " + column.getFieldName(), e);
            }
        }

        // 分页
        if (limit != null && limit > 0) {
            sql.OFFSET("#{offset}").LIMIT("#{limit}");
        }

        return sql.toString();
    }

    // ========== 辅助方法 ==========

    private String getSelectColumns(EntityMetaCache.EntityMeta meta) {
        StringBuilder columns = new StringBuilder();
        for (EntityMetaCache.ColumnMeta column : meta.getColumnList()) {
            columns.append(column.getColumnName()).append(", ");
        }

        if (columns.length() > 0) {
            columns.setLength(columns.length() - 2); // 移除最后的逗号和空格
        }

        return columns.toString();
    }

    private void addWhereCondition(SQL sql, EntityMetaCache.ColumnMeta column, Object value) {
        if (value instanceof String) {
            // 字符串类型支持模糊查询
            String strValue = (String) value;
            if (strValue.contains("%")) {
                sql.WHERE(column.getColumnName() + " LIKE #{" + column.getFieldName() + "}");
            } else {
                sql.WHERE(column.getColumnName() + " = #{" + column.getFieldName() + "}");
            }
        } else if (value instanceof Collection) {
            // 集合类型使用IN查询
            sql.WHERE(column.getColumnName() + " IN (#{" + column.getFieldName() + "})");
        } else {
            sql.WHERE(column.getColumnName() + " = #{" + column.getFieldName() + "}");
        }
    }

    private void addOrderBy(SQL sql, Map<String, Object> params) {
        @SuppressWarnings("unchecked")
        List<String> orderBy = (List<String>) params.get("orderBy");

        if (orderBy != null && !orderBy.isEmpty()) {
            for (String order : orderBy) {
                sql.ORDER_BY(order);
            }
        }
    }

    private void addPagination(SQL sql, Map<String, Object> params) {
        Integer offset = (Integer) params.get("offset");
        Integer limit = (Integer) params.get("limit");

        if (limit != null && limit > 0) {
            if (offset != null && offset > 0) {
                sql.OFFSET("#{offset}").LIMIT("#{limit}");
            } else {
                sql.LIMIT("#{limit}");
            }
        }
    }
}
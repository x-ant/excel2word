package com.xant.component.jdbc.plus;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.util.StrUtil;
import com.xant.common.constant.BaseSqlConstant;
import com.xant.common.util.GenericUtil;
import com.xant.common.util.UUIDUtil;
import org.apache.ibatis.builder.annotation.ProviderContext;
import org.apache.ibatis.jdbc.SQL;

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;

import static cn.hutool.json.XMLTokener.entity;

/**
 * 通用SQL提供器
 *
 * @author xuhq
 */
public class BaseSqlProvider {

    /**
     * 元数据缓存
     */
    private static final Map<Class<?>, EntityMetaCache.EntityMeta> mapperClass2MetaCache = new ConcurrentHashMap<>();

    /**
     * 通用的 selectById 方法
     *
     * @param params  查询参数
     * @param context sql执行上下文
     * @return SQL
     */
    public String selectById(Map<String, Object> params, ProviderContext context) {
        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);
        if (entityMeta.getIdColumn() == null) {
            throw new RuntimeException("实体类没有定义主键字段");
        }

        SQL sql = new SQL()
                .SELECT(getSelectColumns(entityMeta))
                .FROM(entityMeta.getTableName())
                .WHERE(entityMeta.getIdColumn().getColumnName() + " = #{id}");

        return sql.toString();
    }

    /**
     * 通用的 selectById 方法
     *
     * @param params  查询参数
     * @param context sql执行上下文
     * @return SQL
     */
    public String selectByIdList(Map<String, Object> params, ProviderContext context) {
        Collection<String> idList = (Collection<String>) params.get(BaseSqlConstant.ID_LIST);
        if (CollUtil.isEmpty(idList)) {
            throw new RuntimeException("批量ID查询不能为空");
        }
        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);
        if (entityMeta.getIdColumn() == null) {
            throw new RuntimeException("实体类没有定义主键字段");
        }

        // todo 后续处理999问题
        SQL sql = new SQL()
                .SELECT(getSelectColumns(entityMeta))
                .FROM(entityMeta.getTableName())
                .WHERE(entityMeta.getIdColumn().getColumnName() + " IN (#{idList})");

        return sql.toString();
    }

    /**
     * 通用的INSERT方法
     *
     * @param params  查询参数
     * @param context sql执行上下文
     * @return SQL
     */
    public String insert(Map<String, Object> params, ProviderContext context) {
        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);
        SQL sql = new SQL().INSERT_INTO(entityMeta.getTableName());

        for (EntityMetaCache.ColumnMeta column : entityMeta.getColumnList()) {
            // 跳过非插入字段
            if (!column.isInsertable()) {
                continue;
            }

            // 主键
            if (column.isId()) {
                column.setFieldValue(entity, UUIDUtil.getUUID());
            }

            sql.VALUES(column.getColumnName(), "#{" + BaseSqlConstant.ENTITY + StrUtil.DOT + column.getFieldName() + "}");
        }

        return sql.toString();
    }

    /**
     * 通用的 updateById 方法
     *
     * @param params  查询参数
     * @param context sql执行上下文
     * @return SQL
     */
    public String updateById(Map<String, Object> params, ProviderContext context) {
        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);
        if (entityMeta.getIdColumn() == null) {
            throw new RuntimeException("实体类没有定义主键字段");
        }

        SQL sql = new SQL().UPDATE(entityMeta.getTableName());

        for (EntityMetaCache.ColumnMeta column : entityMeta.getColumnList()) {
            // 跳过主键和非更新字段
            if (column.isId() || !column.isUpdatable()) {
                continue;
            }

            sql.SET(column.getColumnName() + " = #{" + BaseSqlConstant.ENTITY + StrUtil.DOT + column.getFieldName() + "}");
        }

        sql.WHERE(entityMeta.getIdColumn().getColumnName() + " = #{ " + BaseSqlConstant.ENTITY + StrUtil.DOT + "id}");

        return sql.toString();
    }

    /**
     * 通用的 deleteById 方法
     *
     * @param params  查询参数
     * @param context sql执行上下文
     * @return SQL
     */
    public String deleteById(Map<String, Object> params, ProviderContext context) {
        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);
        if (entityMeta.getIdColumn() == null) {
            throw new RuntimeException("实体类没有定义主键字段");
        }

        return new SQL()
                .DELETE_FROM(entityMeta.getTableName())
                .WHERE(entityMeta.getIdColumn().getColumnName() + " = #{id}")
                .toString();
    }

    /**
     * 动态条件查询
     *
     * @param params  查询参数
     * @param context sql执行上下文
     * @return SQL
     */
    public String selectByMap(Map<String, Object> params, ProviderContext context) {
        Map<String, Object> fieldMap = (Map<String, Object>) params.get(BaseSqlConstant.FIELD_MAP);
        if (CollUtil.isEmpty(fieldMap)) {
            throw new RuntimeException("查询条件不能为空");
        }

        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);
        SQL sql = new SQL()
                .SELECT(getSelectColumns(entityMeta))
                .FROM(entityMeta.getTableName());

        // 动态WHERE条件
        for (Map.Entry<String, Object> field : fieldMap.entrySet()) {
            EntityMetaCache.ColumnMeta columnMeta = entityMeta.getColumn(field.getKey());
            if (Objects.isNull(columnMeta)) {
                continue;
            }
            addWhereCondition(sql, columnMeta, field.getValue());
        }

        // 排序
        addOrderBy(sql, fieldMap);

        // 分页
        addPagination(sql, fieldMap);

        return sql.toString();
    }

    /**
     * 动态条件删除
     *
     * @param params  查询参数
     * @param context sql执行上下文
     * @return SQL
     */
    public String deleteByMap(Map<String, Object> params, ProviderContext context) {
        Map<String, Object> fieldMap = (Map<String, Object>) params.get(BaseSqlConstant.FIELD_MAP);
        if (CollUtil.isEmpty(fieldMap)) {
            throw new RuntimeException("查询条件不能为空");
        }

        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);
        SQL sql = new SQL()
                .DELETE_FROM(entityMeta.getTableName());

        // 动态WHERE条件
        for (Map.Entry<String, Object> field : fieldMap.entrySet()) {
            EntityMetaCache.ColumnMeta columnMeta = entityMeta.getColumn(field.getKey());
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
     * @param params  查询参数
     * @param context sql执行上下文
     * @return SQL
     */
    public String selectByEntity(Map<String, Object> params, ProviderContext context) {
        Object entity = params.get(BaseSqlConstant.ENTITY);
        if (entity == null) {
            throw new RuntimeException("查询条件不能为空");
        }

        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);
        SQL sql = new SQL()
                .SELECT(getSelectColumns(entityMeta))
                .FROM(entityMeta.getTableName());

        // 动态WHERE条件
        for (EntityMetaCache.ColumnMeta column : entityMeta.getColumnList()) {
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
    public String deleteByIdList(Map<String, Object> params, ProviderContext context) {
        Collection<String> idList = (Collection<String>) params.get(BaseSqlConstant.ID_LIST);
        if (CollUtil.isEmpty(idList)) {
            throw new RuntimeException("批量删除ID不能为空");
        }

        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);

        SQL sql = new SQL()
                .DELETE_FROM(entityMeta.getTableName())
                .WHERE(entityMeta.getIdColumn().getColumnName() + " IN (#{idList})");

        return sql.toString();
    }

    /**
     * 分页查询
     *
     * @param params 查询参数
     * @return SQL
     */
    public String selectPage(Map<String, Object> params, ProviderContext context) {
        Object entity = params.get(BaseSqlConstant.ENTITY);
        if (entity == null) {
            throw new RuntimeException("查询条件不能为空");
        }

        EntityMetaCache.EntityMeta entityMeta = getEntityMeta(context);
        SQL sql = new SQL()
                .SELECT(getSelectColumns(entityMeta))
                .FROM(entityMeta.getTableName());

        // WHERE条件
        for (EntityMetaCache.ColumnMeta column : entityMeta.getColumnList()) {
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
        addPagination(sql, params);

        return sql.toString();
    }

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
                sql.WHERE(column.getColumnName() + " LIKE #{" + BaseSqlConstant.FIELD_MAP + StrUtil.DOT + column.getFieldName() + "}");
            } else {
                sql.WHERE(column.getColumnName() + " = #{" + BaseSqlConstant.FIELD_MAP + StrUtil.DOT + column.getFieldName() + "}");
            }
        } else if (value instanceof Collection) {
            // 集合类型使用IN查询
            sql.WHERE(column.getColumnName() + " IN (#{" + BaseSqlConstant.FIELD_MAP + StrUtil.DOT + column.getFieldName() + "})");
        } else {
            sql.WHERE(column.getColumnName() + " = #{" + BaseSqlConstant.FIELD_MAP + StrUtil.DOT + column.getFieldName() + "}");
        }
    }

    private void addOrderBy(SQL sql, Map<String, Object> params) {
        List<String> orderByList = (List<String>) params.get(BaseSqlConstant.ORDER_BY_LIST);
        if (CollUtil.isEmpty(orderByList)) {
            return;
        }
        for (String orderBy : orderByList) {
            sql.ORDER_BY(orderBy);
        }
    }

    private void addPagination(SQL sql, Map<String, Object> params) {
        Integer offset = (Integer) params.get(BaseSqlConstant.OFFSET);
        Integer limit = (Integer) params.get(BaseSqlConstant.LIMIT);

        if (Objects.nonNull(limit) && limit > 0) {
            if (Objects.nonNull(offset) && offset > 0) {
                sql.OFFSET("#{offset}").LIMIT("#{limit}");
            } else {
                sql.LIMIT("#{limit}");
            }
        }
    }

    private EntityMetaCache.EntityMeta getEntityMeta(ProviderContext context) {
        Class<?> mapperClass = context.getMapperType();
        return mapperClass2MetaCache.computeIfAbsent(mapperClass, k -> {
            Class<?> entityClass = GenericUtil.getInterfacesGenericType(mapperClass, 0);
            return EntityMetaCache.getEntityMeta(entityClass);
        });
    }
}
package com.xant.component.jdbc.plus;

import jakarta.persistence.Column;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import jakarta.persistence.Transient;
import lombok.Getter;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 实体元数据工具类
 *
 * @author xuhq
 */
public class EntityMetaCache {

    private static final Map<Class<?>, EntityMeta> META_CACHE = new ConcurrentHashMap<>();

    /**
     * 实体元数据
     */
    public static class EntityMeta {

        @Getter
        private String tableName;

        @Getter
        private final List<ColumnMeta> columnList = new ArrayList<>();

        @Getter
        private ColumnMeta idColumn;

        private final Map<String, ColumnMeta> columnMap = new HashMap<>();

        public ColumnMeta getColumn(String fieldName) {
            return columnMap.get(fieldName);
        }
    }

    /**
     * 列元数据
     */
    @Getter
    public static class ColumnMeta {

        private Field field;

        private String fieldName;

        private String columnName;

        private boolean isId;

        private boolean insertable;

        private boolean updatable;

        private boolean nullable;

        private Class<?> fieldType;

        public Object getFieldValue(Object entity) {
            try {
                if (!field.isAccessible()) {
                    field.setAccessible(true);
                }
                return field.get(entity);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }

        public void setFieldValue(Object entity, Object fieldValue) {
            try {
                if (!field.isAccessible()) {
                    field.setAccessible(true);
                }
                field.set(entity, fieldValue);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }

    /**
     * 获取实体元数据
     */
    public static EntityMeta getEntityMeta(Class<?> entityClass) {
        return META_CACHE.computeIfAbsent(entityClass, EntityMetaCache::parseEntityMeta);
    }

    /**
     * 解析实体元数据
     *
     * @param entityClass 实体类
     * @return 实体元数据
     */
    private static EntityMeta parseEntityMeta(Class<?> entityClass) {
        EntityMeta meta = new EntityMeta();

        // 获取表名
        Table table = entityClass.getAnnotation(Table.class);
        if (table != null) {
            meta.tableName = table.name();
        } else {
            // 默认表名：类名驼峰转下划线
            meta.tableName = camelToUnderscore(entityClass.getSimpleName());
        }

        // 获取所有字段（包括父类）
        List<Field> allFieldList = getAllFieldList(entityClass);

        for (Field field : allFieldList) {
            // 跳过静态字段
            if (Modifier.isStatic(field.getModifiers())) {
                continue;
            }

            // 跳过非序列化字段
            Transient trans = field.getAnnotation(Transient.class);
            if (trans != null) {
                continue;
            }

            ColumnMeta columnMeta = parseColumnMeta(field);
            meta.columnList.add(columnMeta);
            meta.columnMap.put(columnMeta.fieldName, columnMeta);

            // 记录主键字段
            if (columnMeta.isId) {
                meta.idColumn = columnMeta;
            }
        }

        return meta;
    }

    /**
     * 解析字段元数据
     *
     * @param field 字段定义
     * @return 字段元数据
     */
    private static ColumnMeta parseColumnMeta(Field field) {
        ColumnMeta meta = new ColumnMeta();
        meta.field = field;
        meta.fieldName = field.getName();
        meta.fieldType = field.getType();

        // 检查是否为主键
        meta.isId = Objects.nonNull(field.getAnnotation(Id.class));

        // 获取列注解
        Column column = field.getAnnotation(Column.class);
        if (column != null) {
            meta.columnName = column.name().isEmpty() ?
                    camelToUnderscore(field.getName()) : column.name();
            meta.insertable = column.insertable();
            meta.updatable = column.updatable();
            meta.nullable = column.nullable();
        } else {
            // 默认列名
            meta.columnName = camelToUnderscore(field.getName());
            meta.insertable = true;
            meta.updatable = true;
            meta.nullable = true;
        }

        return meta;
    }

    /**
     * 获取所有字段（包括父类）
     *
     * @param clazz 类定义
     */
    private static List<Field> getAllFieldList(Class<?> clazz) {
        List<Field> fields = new ArrayList<>();
        while (Objects.nonNull(clazz) && clazz != Object.class) {
            fields.addAll(Arrays.asList(clazz.getDeclaredFields()));
            clazz = clazz.getSuperclass();
        }
        return fields;
    }

    /**
     * 驼峰转下划线
     */
    public static String camelToUnderscore(String camel) {
        if (camel == null || camel.isEmpty()) {
            return camel;
        }

        StringBuilder result = new StringBuilder();
        result.append(Character.toLowerCase(camel.charAt(0)));

        for (int i = 1; i < camel.length(); i++) {
            char ch = camel.charAt(i);
            if (Character.isUpperCase(ch)) {
                result.append('_');
                result.append(Character.toLowerCase(ch));
            } else {
                result.append(ch);
            }
        }

        return result.toString();
    }

}
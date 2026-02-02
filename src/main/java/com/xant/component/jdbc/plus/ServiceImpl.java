package com.xant.component.jdbc.plus;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.lang.Assert;
import cn.hutool.core.util.ObjectUtil;
import cn.hutool.core.util.StrUtil;
import com.xant.component.jdbc.SqlSessionFactorySingleton;
import com.xant.component.jdbc.TransactionContextManager;
import com.xant.component.jdbc.TransactionUtil;
import com.xant.common.util.GenericUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.ibatis.reflection.ExceptionUtil;
import org.apache.ibatis.session.ExecutorType;
import org.apache.ibatis.session.SqlSession;
import org.slf4j.Logger;

import java.util.Collection;
import java.util.Objects;
import java.util.function.BiConsumer;
import java.util.function.BiPredicate;
import java.util.function.Consumer;

@Slf4j
public abstract class ServiceImpl<M extends BaseMapper<T>, T> implements IService<T> {

    @Override
    public M getBaseMapper() {
        return (M) SqlSessionFactorySingleton.getSingleSqlSession().getMapper(mapperClass);
    }

    protected Class<?> entityClass = currentModelClass();

    protected Class<?> mapperClass = currentMapperClass();

    protected Class<T> currentMapperClass() {
        return (Class<T>) GenericUtil.getSuperClassGenericType(getClass(), 0);
    }

    protected Class<T> currentModelClass() {
        return (Class<T>) GenericUtil.getSuperClassGenericType(getClass(), 1);
    }

    /**
     * 批量插入
     *
     * @param entityList ignore
     * @param batchSize  ignore
     * @return ignore
     */
    @Override
    public boolean saveBatch(Collection<T> entityList, int batchSize) {
        return TransactionUtil.transactionWithRequired(() -> {
            String sqlStatement = getSqlStatementId("insert");
            return executeBatch(entityList, batchSize, (sqlSession, entity) -> sqlSession.insert(sqlStatement, entity));
        });
    }

    /**
     * TableId 注解存在更新记录，否插入一条记录
     *
     * @param entity 实体对象
     * @return boolean
     */
    @Override
    public boolean saveOrUpdate(T entity) {
        if (Objects.nonNull(entity)) {
            return TransactionUtil.transactionWithRequired(() -> {
                EntityMetaCache.EntityMeta entityMeta = EntityMetaCache.getEntityMeta(entityClass);
                EntityMetaCache.ColumnMeta idColumn = entityMeta.getIdColumn();
                Object idVal = idColumn.getFieldValue(entity);
                return ObjectUtil.isEmpty(idVal) ? save(entity) : updateById(entity);
            });
        }
        return false;
    }

    @Override
    public boolean saveOrUpdateBatch(Collection<T> entityList, int batchSize) {
        EntityMetaCache.EntityMeta entityMeta = EntityMetaCache.getEntityMeta(entityClass);
        EntityMetaCache.ColumnMeta idColumn = entityMeta.getIdColumn();
        return TransactionUtil.transactionWithRequired(() -> {
            return saveOrUpdateBatch(log, entityList, batchSize, (sqlSession, entity) -> {
                Object idVal = idColumn.getFieldValue(entity);
                return ObjectUtil.isEmpty(idVal)
                        || CollectionUtils.isEmpty(sqlSession.selectList(getSqlStatementId("selectById"), entity));
            }, (sqlSession, entity) -> {
                sqlSession.update(getSqlStatementId("updateById"), entity);
            });
        });
    }

    @Override
    public boolean updateBatchById(Collection<T> entityList, int batchSize) {
        return TransactionUtil.transactionWithRequired(() -> {
            String sqlStatement = getSqlStatementId("updateById");
            return executeBatch(entityList, batchSize, (sqlSession, entity) -> {
                sqlSession.update(sqlStatement, entity);
            });
        });
    }

    /**
     * 执行批量操作
     *
     * @param list      数据集合
     * @param batchSize 批量大小
     * @param consumer  执行方法
     * @param <E>       泛型
     * @return 操作结果
     * @since 3.3.1
     */
    protected <E> boolean executeBatch(Collection<E> list, int batchSize, BiConsumer<SqlSession, E> consumer) {
        return executeBatch(this.entityClass, log, list, batchSize, consumer);
    }

    /**
     * 执行批量操作
     *
     * @param entityClass 实体类
     * @param log         日志对象
     * @param list        数据集合
     * @param batchSize   批次大小
     * @param consumer    consumer
     * @param <E>         T
     * @return 操作结果
     */
    private <E> boolean executeBatch(Class<?> entityClass, Logger log, Collection<E> list, int batchSize, BiConsumer<SqlSession, E> consumer) {
        Assert.isFalse(batchSize < 1, "batchSize must not be less than one");
        return !CollUtil.isEmpty(list) && executeBatch(entityClass, log, sqlSession -> {
            int size = list.size();
            int i = 1;
            for (E element : list) {
                consumer.accept(sqlSession, element);
                if ((i % batchSize == 0) || i == size) {
                    sqlSession.flushStatements();
                }
                i++;
            }
        });
    }

    /**
     * 执行批量操作
     *
     * @param entityClass 实体
     * @param log         日志对象
     * @param consumer    consumer
     * @return 操作结果
     */
    private boolean executeBatch(Class<?> entityClass, Logger log, Consumer<SqlSession> consumer) {
        boolean transaction = TransactionContextManager.getInTransaction();
        SqlSession originSqlSession = TransactionContextManager.getSqlSession();
        if (Objects.nonNull(originSqlSession)) {
            //原生无法支持执行器切换，当存在批量操作时，会嵌套两个session的，优先commit上一个session
            //按道理来说，这里的值应该一直为false。
            originSqlSession.commit(!transaction);
        }
        SqlSession sqlSession = SqlSessionFactorySingleton.getSqlSessionFactory().openSession(ExecutorType.BATCH);
        if (!transaction) {
            log.warn("SqlSession [" + sqlSession + "] was not registered for synchronization because DataSource is not transactional");
        }
        try {
            consumer.accept(sqlSession);
            //非事物情况下，强制commit。
            sqlSession.commit(!transaction);
            return true;
        } catch (Throwable t) {
            sqlSession.rollback();
            throw new RuntimeException(ExceptionUtil.unwrapThrowable(t));
        } finally {
            sqlSession.close();
        }
    }

    /**
     * 批量更新或保存
     *
     * @param log       日志对象
     * @param list      数据集合
     * @param batchSize 批次大小
     * @param predicate predicate(新增条件) notNull
     * @param consumer  consumer（更新处理） notNull
     * @param <E>       E
     * @return 操作结果
     */
    private <E> boolean saveOrUpdateBatch(Logger log, Collection<E> list, int batchSize, BiPredicate<SqlSession, E> predicate, BiConsumer<SqlSession, E> consumer) {
        String sqlStatement = getSqlStatementId("insert");
        return executeBatch(entityClass, log, list, batchSize, (sqlSession, entity) -> {
            if (predicate.test(sqlSession, entity)) {
                sqlSession.insert(sqlStatement, entity);
            } else {
                consumer.accept(sqlSession, entity);
            }
        });
    }

    /**
     * 获取mapperStatementId
     *
     * @param sqlMethod 方法名
     * @return 命名id
     */
    private String getSqlStatementId(String sqlMethod) {
        return mapperClass.getName() + StrUtil.DOT + sqlMethod;
    }

}
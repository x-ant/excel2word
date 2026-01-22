package com.xant.component.jdbc;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.SQLException;

/**
 * 将数据库连接与线程绑定，得到一个线程，一个数据库，对应一个连接
 * 通过静态方法实现工具属性
 *
 * @author xuhq0808
 */
public class SingleThreadConnectionHolder {

    /**
     * 每个线程一个ConnectionHolder
     */
    private static ThreadLocal<ConnectionHolder> connectionThreadLocal = new ThreadLocal<>();

    /**
     * 获取当前线程对应的ConnectionHolder
     *
     * @return 当前线程对应的ConnectionHolder
     */
    private static ConnectionHolder getConnectionHolder() {
        ConnectionHolder connectionHolder = connectionThreadLocal.get();

        if (connectionHolder == null) {
            connectionHolder = new ConnectionHolder();
            connectionThreadLocal.set(connectionHolder);
        }

        return connectionHolder;
    }

    /**
     * 通过数据库实例对象，获取当前线程对应的Connection
     *
     * @param dataSource 数据库实例
     * @return 对应的数据库连接
     * @throws SQLException 数据库操作异常
     */
    public static Connection getConnection(DataSource dataSource) throws SQLException {
        return getConnectionHolder().getConnectionByDataSource(dataSource);
    }
}

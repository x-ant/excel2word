package com.xant.component.jdbc;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.Map;

/**
 * 从数据库实例中获取数据库连接，如果没有则创建，只会保存第一条创建的数据库连接
 *
 * @author xuhq
 */
public class ConnectionHolder {

    /**
     * 每一个数据源对应一个连接，放到ThreadLocal里则变成
     * 每一个线程每一个数据源对应一个连接
     */
    private Map<DataSource, Connection> dataSourceConnectionMap = new HashMap<>();

    /**
     * 根据数据库实例获取数据库连接，如果没有则创建
     *
     * @param dataSource 数据库实例
     * @return 对应类型数据库连接
     * @throws SQLException 数据库连接操作异常
     */
    public Connection getConnectionByDataSource(DataSource dataSource) throws SQLException {
        Connection conn = dataSourceConnectionMap.get(dataSource);

        if (conn == null || conn.isClosed()) {
            conn = dataSource.getConnection();
            dataSourceConnectionMap.put(dataSource, conn);
        }

        return conn;
    }

}

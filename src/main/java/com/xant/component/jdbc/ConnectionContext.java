package com.xant.component.jdbc;

import com.xant.component.ThreadContext;

import javax.sql.DataSource;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

/**
 * 将数据库连接与线程绑定，得到一个线程，一个数据库，对应一个连接
 *
 * @author xuhq0808
 */
public class ConnectionContext {

    private static final String CONNECTION_HOLDER_KEY = "CONNECTION_HOLDER";

    public static Connection getConnection(DataSource dataSource) throws SQLException {
        return getConnectionHolder().getConnection(dataSource);
    }

    public static void removeConnection(DataSource dataSource) {
        getConnectionHolder().removeConnection(dataSource);
    }

    private ConnectionContext() {
    }

    private static ConnectionHolder getConnectionHolder() {
        ConnectionHolder connectionHolder = ThreadContext.getAttribute(CONNECTION_HOLDER_KEY);

        if (connectionHolder == null) {
            connectionHolder = new ConnectionHolder();
            ThreadContext.setAttribute(CONNECTION_HOLDER_KEY, connectionHolder);
        }

        return connectionHolder;
    }

    private static class ConnectionHolder {

        /**
         * 每一个数据源对应一个连接，放到ThreadLocal里则变成
         * 每一个线程每一个数据源对应一个连接
         */
        private final Map<DataSource, Connection> dataSourceConnectionMap = new HashMap<>();

        public Connection getConnection(DataSource dataSource) throws SQLException {
            Connection conn = dataSourceConnectionMap.get(dataSource);

            if (conn == null || conn.isClosed()) {
                conn = dataSource.getConnection();
                dataSourceConnectionMap.put(dataSource, conn);
            }

            return conn;
        }

        public void removeConnection(DataSource dataSource) {
            Connection conn = dataSourceConnectionMap.get(dataSource);

            if (Objects.nonNull(conn)) {
                dataSourceConnectionMap.remove(dataSource);
            }
        }

    }
}
package com.xant.component.jdbc;

import com.zaxxer.hikari.HikariConfig;
import com.zaxxer.hikari.HikariDataSource;

import javax.sql.DataSource;

/**
 * sqlite数据源单例
 *
 * @author xuhq
 */
public class SqliteSingleton {

    private static volatile DataSource instance;

    private SqliteSingleton() {
    }

    public static DataSource getSingleton() {
        if (instance == null) {
            synchronized (SqliteSingleton.class) {
                if (instance == null) {
                    HikariConfig config = new HikariConfig();

                    config.setJdbcUrl("jdbc:sqlite:data.db");
                    config.setDriverClassName("org.sqlite.JDBC");
                    config.setMaximumPoolSize(3);
                    config.setMinimumIdle(0);
                    config.setConnectionTimeout(30000);
                    config.setIdleTimeout(600000);

                    config.addDataSourceProperty("journal_mode", "WAL");
                    config.addDataSourceProperty("synchronous", "NORMAL");
                    config.addDataSourceProperty("cache_size", "-20000");
                    config.addDataSourceProperty("busy_timeout", "5000");
                    instance = new HikariDataSource(config);
                }
            }
        }
        return instance;
    }
}

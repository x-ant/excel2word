package com.xant.component.jdbc;

import com.zaxxer.hikari.HikariConfig;
import com.zaxxer.hikari.HikariDataSource;
import org.apache.ibatis.datasource.DataSourceFactory;

import javax.sql.DataSource;
import java.util.Properties;

/**
 * HikariCP数据源工厂
 *
 * @author xuhq
 */
public class HikariCPDataSourceFactory implements DataSourceFactory {

    private DataSource dataSource;

    @Override
    public void setProperties(Properties props) {
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
        this.dataSource = new HikariDataSource(config);
    }

    @Override
    public DataSource getDataSource() {
        return dataSource;
    }
}

package com.xant.component.jdbc;

import org.apache.ibatis.datasource.DataSourceFactory;

import javax.sql.DataSource;
import java.util.Properties;

/**
 * HikariCP数据源工厂
 *
 * @author xuhq
 */
public class HikariCPDataSourceFactory implements DataSourceFactory {

    @Override
    public void setProperties(Properties props) {
    }

    @Override
    public DataSource getDataSource() {
        return SqliteSingleton.getSingleton();
    }
}

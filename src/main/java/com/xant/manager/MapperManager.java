package com.xant.manager;

import com.xant.component.SqlSessionManagerSingleton;
import com.xant.dao.BaseMapper;
import lombok.extern.slf4j.Slf4j;
import org.apache.ibatis.session.SqlSession;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 数据库操作管理
 *
 * @author xuhq
 */
@Slf4j
public class MapperManager {

    private static final Map<Class<?>, Object> mapperProxyMap = new ConcurrentHashMap<>();

    private MapperManager() {
    }

    public static <P, T extends BaseMapper<P>> T getMapper(Class<T> clazz) {
        SqlSession sqlSession = SqlSessionManagerSingleton.getSingleton();
        Object mapperProxy = mapperProxyMap.computeIfAbsent(clazz, k -> {
            T originProxyMapper = sqlSession.getMapper(clazz);
            return Proxy.newProxyInstance(clazz.getClassLoader(), clazz.getInterfaces(), new ProxyInvocationHandler<>(originProxyMapper));
        });
        return clazz.cast(mapperProxy);
    }

    private static class ProxyInvocationHandler<P> implements InvocationHandler {

        private final BaseMapper<P> baseMapper;

        public <T extends BaseMapper<P>> ProxyInvocationHandler(BaseMapper<P> baseMapper) {
            this.baseMapper = baseMapper;
        }

        @Override
        public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
            if (Object.class.equals(method.getDeclaringClass())) {
                return method.invoke(this, args);
            }
            if (method.getName().equals("selectById")) {
                return baseMapper.selectById((String) args[0]);
            }
            if (method.getName().equals("insert")) {
                return baseMapper.insert((P) args[0]);
            }
            if (method.getName().equals("updateById")) {
                return baseMapper.updateById((P) args[0]);
            }
            if (method.getName().equals("deleteById")) {
                return baseMapper.deleteById((String) args[0]);
            }
            return null;
        }

    }

}

package com.xant.component.jdbc;

import java.util.Map;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;

/**
 * 线程上下文
 *
 * @author xuhq
 */
public final class ThreadContext {

    private static final ThreadLocal<Map<Object, Object>> THREAD_LOCAL = new ThreadLocal<>();

    private ThreadContext() {
    }

    public static <T> T getAttribute(Object key) {
        return (T) getAttributeMap().get(key);
    }

    public static void setAttribute(Object key, Object value) {
        getAttributeMap().put(key, value);
    }

    public static <T> T removeAttribute(Object key) {
        return (T) getAttributeMap().remove(key);
    }

    public static boolean containAttribute(Object key) {
        return getAttributeMap().containsKey(key);
    }

    public static void clearAttribute() {
        getAttributeMap().clear();
    }

    private static Map<Object, Object> getAttributeMap() {
        Map<Object, Object> attributeMap = THREAD_LOCAL.get();
        if (Objects.isNull(attributeMap)) {
            attributeMap = new ConcurrentHashMap<>();
            THREAD_LOCAL.set(attributeMap);
        }
        return attributeMap;
    }

}

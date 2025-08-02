package com.xant.component;

import ch.qos.logback.classic.pattern.ClassicConverter;
import ch.qos.logback.classic.spi.ILoggingEvent;

/**
 * 包名简化处理，包名只展示一个字母
 *
 * @author xuhq
 */
public class SimplifiedPackageConverter extends ClassicConverter {

    @Override
    public String convert(ILoggingEvent event) {
        String loggerName = event.getLoggerName();
        return shortenPackageName(loggerName);
    }

    private String shortenPackageName(String fullClassName) {
        if (fullClassName == null) return "";

        int lastDotIndex = fullClassName.lastIndexOf('.');
        if (lastDotIndex == -1) return fullClassName;

        String packagePart = fullClassName.substring(0, lastDotIndex);
        String classPart = fullClassName.substring(lastDotIndex + 1);

        StringBuilder shortened = new StringBuilder();
        for (String pkg : packagePart.split("\\.")) {
            if (!pkg.isEmpty()) {
                shortened.append(pkg.charAt(0)).append('.');
            }
        }
        shortened.append(classPart);

        return shortened.toString();
    }
}
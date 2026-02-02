package com.xant.common.util;

import com.fasterxml.uuid.Generators;
import com.fasterxml.uuid.impl.TimeBasedEpochGenerator;

/**
 * 随时间递增的uuid
 *
 * @author xuhq
 */
public class UUIDUtil {

    private static class GenerateHolder {
        private static final TimeBasedEpochGenerator generator = Generators.timeBasedEpochGenerator();
    }

    public static String getUUID() {
        return GenerateHolder.generator.generate().toString().replaceAll("-", "");
    }

}

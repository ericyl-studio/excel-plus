package com.ericyl.excel.util;

import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Objects;

public class ObjectUtils {

    /**
     * 列坐标 -> 下标
     *
     * @param s
     * @return
     */
    public static int convertToNumber(String s) {
        int result = 0;
        for (int i = 0; i < s.length(); i++) {
            // 将字符转换为0-25之间的数字，然后根据位置加权（26的幂）
            result = result * 26 + (s.charAt(i) - 'A' + 1);
        }
        return result;
    }

    /**
     * 反射方式写入数据
     *
     * @param obj
     * @param field
     * @param value
     * @param <T>
     */
    public static <T> void setField(T obj, Field field, Object value) {
        if (value == null)
            return;
        try {
            field.setAccessible(true);
            field.set(obj, value);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 反射获取数据
     */
    public static Object getField(Object obj, Field field) {
        if (obj == null)
            return null;
        try {
            field.setAccessible(true);
            return field.get(obj);
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 判断数据不为空
     */
    public static <T> boolean isNotEmpty(T t) {
        return !isEmpty(t);
    }

    /**
     * 反射方式判断数据是否为空
     * TODO 尚未支持复杂子属性
     */
    public static <T> boolean isEmpty(T t) {
        if (t == null)
            return true;
        Class<?> clazz = t.getClass();
        Field[] fields = clazz.getDeclaredFields();
        return Arrays.stream(fields).allMatch(field -> {
            field.setAccessible(true);
            try {
                Object obj = field.get(t);
                if (obj == null)
                    return true;
                if (obj instanceof String) {
                    return StringUtils.isEmpty(obj.toString());
                } else if (Number.class.isAssignableFrom(obj.getClass())) {
                    return Objects.equals(new BigDecimal("0.0"), new BigDecimal(String.valueOf(obj)));
                }
                return false;
            } catch (IllegalAccessException e) {
                return true;
            }
        });
    }
}


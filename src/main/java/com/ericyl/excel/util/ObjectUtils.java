package com.ericyl.excel.util;

import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Objects;

/**
 * 对象操作工具类
 * <p>
 * 提供反射操作、类型转换等通用功能
 * </p>
 * 
 * @author ericyl
 * @since 1.0
 */
public class ObjectUtils {

    /**
     * Excel列坐标转换为数字索引
     * <p>
     * 将Excel的列坐标（如 "A", "B", "AA", "AB"）转换为从1开始的数字索引。
     * 转换规则类似26进制，其中 A=1, B=2, ..., Z=26, AA=27, AB=28，以此类推。
     * </p>
     * 
     * @param s Excel列坐标字符串（如 "A", "AA"）
     * @return 对应的数字索引（从1开始）
     * 
     * @example
     * 
     *          <pre>
     * convertToNumber("A")  // 返回 1
     * convertToNumber("B")  // 返回 2
     * convertToNumber("Z")  // 返回 26
     * convertToNumber("AA") // 返回 27
     * convertToNumber("AB") // 返回 28
     *          </pre>
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
     * 通过反射设置对象字段的值
     * <p>
     * 使用反射机制将指定的值设置到对象的字段中。
     * 会自动设置字段的访问权限。
     * </p>
     * 
     * @param obj   目标对象
     * @param field 要设置的字段
     * @param value 要设置的值
     * @param <T>   对象类型
     * @throws RuntimeException 当设置失败时抛出，包装了原始的IllegalAccessException
     */
    public static <T> void setField(T obj, Field field, Object value) {
        if (value == null)
            return;
        try {
            field.setAccessible(true);
            field.set(obj, value);
        } catch (IllegalAccessException e) {
            throw new RuntimeException("设置字段值失败: " + field.getName(), e);
        }
    }

    /**
     * 通过反射获取对象字段的值
     * <p>
     * 使用反射机制获取对象指定字段的值。
     * 会自动设置字段的访问权限。
     * </p>
     * 
     * @param obj   目标对象
     * @param field 要获取的字段
     * @return 字段的值，如果对象为null则返回null
     * @throws RuntimeException 当获取失败时抛出，包装了原始的IllegalAccessException
     */
    public static Object getField(Object obj, Field field) {
        if (obj == null)
            return null;
        try {
            field.setAccessible(true);
            return field.get(obj);
        } catch (IllegalAccessException e) {
            throw new RuntimeException("获取字段值失败: " + field.getName(), e);
        }
    }

    /**
     * 判断对象是否不为空
     * <p>
     * 与 {@link #isEmpty(Object)} 相反
     * </p>
     * 
     * @param t   要判断的对象
     * @param <T> 对象类型
     * @return 如果对象不为空返回true，否则返回false
     */
    public static <T> boolean isNotEmpty(T t) {
        return !isEmpty(t);
    }

    /**
     * 判断对象是否为空
     * <p>
     * 通过反射检查对象的所有字段，判断对象是否为"空"。
     * 判断规则：
     * <ul>
     * <li>对象本身为null，返回true</li>
     * <li>所有字段都为null，返回true</li>
     * <li>String类型字段为空字符串，视为空</li>
     * <li>Number类型字段值为0，视为空</li>
     * <li>其他类型字段为null，视为空</li>
     * </ul>
     * 注意：尚未支持复杂子对象的递归判断
     * </p>
     * 
     * @param t   要判断的对象
     * @param <T> 对象类型
     * @return 如果对象为空返回true，否则返回false
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

                // 字符串类型：空字符串视为空
                if (obj instanceof String) {
                    return StringUtils.isEmpty(obj.toString());
                }
                // 数字类型：0值视为空
                else if (Number.class.isAssignableFrom(obj.getClass())) {
                    return Objects.equals(new BigDecimal("0.0"), new BigDecimal(String.valueOf(obj)));
                }
                // 其他类型：非null即不为空
                return false;
            } catch (IllegalAccessException e) {
                // 访问失败时视为空
                return true;
            }
        });
    }
}

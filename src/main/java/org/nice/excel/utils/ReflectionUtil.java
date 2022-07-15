package org.nice.excel.utils;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;

/**
 * <p>
 * 反射工具类
 * </p>
 *
 * @author WECENG
 * @since 2021/7/6 15:49
 */
public class ReflectionUtil {

    /**
     * 获取泛型类类型
     *
     * @param clazz 泛型类
     * @return
     */
    public static Class<?> getGenericsClass(Class<?> clazz) {
        Type[] typeArguments = ((ParameterizedType) clazz.getGenericSuperclass()).getActualTypeArguments();
        return findActualType(typeArguments);
    }

    /**
     * 获取类型
     *
     * @param typeArguments 类型数组
     * @return
     */
    private static Class<?> findActualType(Type[] typeArguments) {
        for (Type type : typeArguments) {
            if (type instanceof Class) {
                return (Class<?>) type;
            }
            if (type instanceof ParameterizedType) {
                Type[] actualTypeArguments = ((ParameterizedType) type).getActualTypeArguments();
                return findActualType(actualTypeArguments);
            }
        }
        return null;
    }

}

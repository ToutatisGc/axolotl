package cn.toutatis.xvoid.axolotl.toolkit;

import java.lang.reflect.Field;

/**
 * 字段工具类
 * @author 张智凯
 * @version 1.0
 */
public class FieldToolkit {

    /**
     * 递归获取字段，包含父级。
     * @param clazz 类
     * @param fieldName 字段名
     * @return 字段
     */
    public static Field recursionGetField(Class<?> clazz, String fieldName){
        return recursionGetField(clazz,fieldName,true);
    }

    /**
     * 递归获取字段，包含父级。
     * @param clazz 类
     * @param fieldName 字段名
     * @param callSuperClass 是否调用父类
     * @return 字段
     */
    public static Field recursionGetField(Class<?> clazz, String fieldName, boolean callSuperClass){
        try {
            return clazz.getDeclaredField(fieldName);
        } catch (NoSuchFieldException e) {
            if (clazz.getSuperclass() != null && callSuperClass){
                return recursionGetField(clazz.getSuperclass(),fieldName);
            }
            return null;
        }
    }
}

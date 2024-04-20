package cn.toutatis.xvoid.axolotl.toolkit;

import java.lang.reflect.Field;

/**
 * @author 张智凯
 * @version 1.0
 * @data 2024/4/20 17:28
 */
public class FieldToolkit {

    public static Field recursionGetField(Class<?> clazz, String fieldName){
        try {
            return clazz.getDeclaredField(fieldName);
        } catch (NoSuchFieldException e) {
            if (clazz.getSuperclass() != null){
                return recursionGetField(clazz.getSuperclass(),fieldName);
            }
            return null;
        }
    }
}

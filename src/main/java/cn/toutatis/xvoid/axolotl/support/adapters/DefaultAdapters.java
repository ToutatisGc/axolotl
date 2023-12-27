package cn.toutatis.xvoid.axolotl.support.adapters;

import cn.toutatis.xvoid.axolotl.support.DataCastAdapter;

import java.util.Map;

/**
 * 默认数据类型适配器
 * @author Toutatis_Gc
 */
public class DefaultAdapters {

    private static final Map<Class<?>,DataCastAdapter<?>> defaultAdapters;

    static {
        defaultAdapters = Map.of(
                String.class, new DefaultStringAdapter()
//                Integer.class, new DefaultIntegerAdapter(),
//                Long.class, new DefaultLongAdapter(),
//                Double.class, new DefaultDoubleAdapter(),
//                Float.class, new DefaultFloatAdapter(),
//                Boolean.class, new DefaultBooleanAdapter(),
//                Short.class, new DefaultShortAdapter(),
//                Byte.class, new DefaultByteAdapter(),
//                Character.class, new DefaultCharacterAdapter(),
        );
    }

    public static DataCastAdapter<?> getAdapter(Class<?> clazz) {
        return defaultAdapters.get(clazz);
    }

}

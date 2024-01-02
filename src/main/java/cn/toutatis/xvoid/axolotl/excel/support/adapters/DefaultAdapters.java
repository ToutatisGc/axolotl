package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;

import java.util.Map;

/**
 * 默认数据类型适配器
 * @author Toutatis_Gc
 */
public class DefaultAdapters {

    private static final Map<Class<?>,DataCastAdapter<?>> defaultAdapters;

    static {
        defaultAdapters = Map.of(
                String.class, new DefaultStringAdapter(),
                Integer.class, new DefaultNumericAdapter<>(Integer.class),
                int.class, new DefaultNumericAdapter<>(Integer.class),
                Long.class, new DefaultNumericAdapter<>(Long.class),
                Double.class, new DefaultNumericAdapter<>(Double.class),
                Float.class, new DefaultNumericAdapter<>(Float.class)
//                Boolean.class, new DefaultBooleanAdapter()
        );
    }

    public static DataCastAdapter<?> getAdapter(Class<?> clazz) {
        return defaultAdapters.get(clazz);
    }

}

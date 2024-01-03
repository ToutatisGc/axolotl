package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;

import java.time.LocalDateTime;
import java.util.Date;
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
                long.class, new DefaultNumericAdapter<>(Long.class),
                Boolean.class, new DefaultBooleanAdapter(),
                boolean.class, new DefaultBooleanAdapter(),
                Date.class, new DefaultDateTimeAdapter<>(Date.class),
                LocalDateTime.class, new DefaultDateTimeAdapter<>(LocalDateTime.class)
        );
    }

    public static DataCastAdapter<?> getAdapter(Class<?> clazz) {
        return defaultAdapters.get(clazz);
    }

}

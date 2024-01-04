package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;

import java.time.LocalDateTime;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * 默认数据类型适配器
 * @author Toutatis_Gc
 */
public class DefaultAdapters {

    /**
     * 默认数据类型适配器
     */
    private static final Map<Class<?>,DataCastAdapter<?>> defaultAdapters = new HashMap<>();

    static {
        defaultAdapters.put(String.class, new DefaultStringAdapter());
        defaultAdapters.put(Integer.class, new DefaultNumericAdapter<>(Integer.class));
        defaultAdapters.put(int.class, new DefaultNumericAdapter<>(Integer.class));
        defaultAdapters.put(Long.class, new DefaultNumericAdapter<>(Long.class));
        defaultAdapters.put(long.class, new DefaultNumericAdapter<>(Long.class));
        defaultAdapters.put(Double.class, new DefaultNumericAdapter<>(Double.class));
        defaultAdapters.put(double.class, new DefaultNumericAdapter<>(Double.class));
        defaultAdapters.put(Boolean.class, new DefaultBooleanAdapter());
        defaultAdapters.put(boolean.class, new DefaultBooleanAdapter());
        defaultAdapters.put(Date.class, new DefaultDateTimeAdapter<>(Date.class));
        defaultAdapters.put(LocalDateTime.class, new DefaultDateTimeAdapter<>(LocalDateTime.class));
    }

    /**
     * 获取默认适配器
     * @param clazz 需要转换的类型
     * @return 数据转换适配器
     */
    public static DataCastAdapter<?> getAdapter(Class<?> clazz) {
        return defaultAdapters.get(clazz);
    }

    /**
     * 注册默认适配器
     * 注意：注册后，会覆盖已有的适配器
     * @param clazz 需要转换的类型
     * @param adapter 数据转换适配器
     */
    public static void registerDefaultAdapter(Class<?> clazz, DataCastAdapter<?> adapter) {
        defaultAdapters.put(clazz, adapter);
    }
}

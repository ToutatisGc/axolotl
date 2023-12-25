package cn.toutatis.xvoid.axolotl.support;

import cn.toutatis.xvoid.axolotl.constant.ReadExcelFeature;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.util.HashMap;
import java.util.Map;

@ToString
@Getter
@Setter
public class WorkBookReaderConfig {

    /**
     * 表索引
     */
    private int sheetIndex = -1;

    /**
     * 表名
     */
    private String sheetName;

    /**
     * 读取的Java类型
     */
    private Class<?> clazz;

    /**
     * 读取的特性
     */
    private Map<ReadExcelFeature, Object> readFeature = new HashMap<>();

    public WorkBookReaderConfig() {
        this(true);
    }
    public WorkBookReaderConfig(boolean withDefaultConfig) {
        if (withDefaultConfig) {
            readFeature.putAll(defaultReadFeature());
        }
    }

    private Map<ReadExcelFeature, Object> defaultReadFeature() {
        Map<ReadExcelFeature, Object> defaultReadFeature = new HashMap<>();
        defaultReadFeature.put(ReadExcelFeature.IGNORE_EMPTY_SHEET_ERROR,true);
        return defaultReadFeature;
    }

    public boolean getReadFeatureAsBoolean(ReadExcelFeature feature) {
        return readFeature.containsKey(feature) && (boolean) readFeature.get(feature);
    }

    public void addReadFeature(ReadExcelFeature feature, Object value) {
        readFeature.put(feature, value);
    }
}

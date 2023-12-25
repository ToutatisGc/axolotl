package cn.toutatis.xvoid.axolotl.support;

import cn.toutatis.xvoid.axolotl.constant.ReadExcelFeature;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

@ToString
@Getter
@Setter
public class WorkBookReaderConfig<T> {

    /**
     * 表索引
     */
    private int sheetIndex = -1;

    /**
     * 表名
     */
    private String sheetName;

    /**
     * 初始行偏移量
     */
    private int initialRowPositionOffset = 0;

    /**
     * 读取的Java类型
     */
    private Class<T> castClass;


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
        defaultReadFeature.put(ReadExcelFeature.SORTED_READ_SHEET_DATA,true);
        defaultReadFeature.put(ReadExcelFeature.INCLUDE_EMPTY_ROW,true);
        defaultReadFeature.put(ReadExcelFeature.USE_MAP_DEBUG,true);
        defaultReadFeature.put(ReadExcelFeature.DATA_BIND_PRECISE_LOCALIZATION,true);
        defaultReadFeature.put(ReadExcelFeature.CAST_NUMBER_TO_DATE,true);
        return defaultReadFeature;
    }

    public boolean getReadFeatureAsBoolean(ReadExcelFeature feature) {
        return readFeature.containsKey(feature) && (boolean) readFeature.get(feature);
    }

    public void addReadFeature(ReadExcelFeature feature, Object value) {
        readFeature.put(feature, value);
    }

    /**
     * 获取转换类型实例
     * @return 由类型转换的生成的实例
     */
    @SuppressWarnings("unchecked")
    public T getCastClassInstance(){
        if(castClass!=null){
            try {
                if (castClass == Map.class){
                    if (getReadFeatureAsBoolean(ReadExcelFeature.SORTED_READ_SHEET_DATA)){
                        return (T) new LinkedHashMap<String,Object>();
                    }else{
                        return (T) new HashMap<String, Object>();
                    }
                }
                return castClass.getDeclaredConstructor().newInstance();
            } catch (InstantiationException | IllegalAccessException |
                     InvocationTargetException | NoSuchMethodException e) {
                throw new RuntimeException(e);
            }
        }else{
            throw new IllegalArgumentException("转换类型为空");
        }
    }
}

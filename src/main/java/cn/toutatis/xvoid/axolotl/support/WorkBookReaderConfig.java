package cn.toutatis.xvoid.axolotl.support;

import cn.toutatis.xvoid.axolotl.constant.ReadExcelFeature;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;

import static cn.toutatis.xvoid.axolotl.constant.ReadExcelFeature.*;

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

    public void setCastClass(Class<T> castClass) {
        this.castClass = castClass;
        Field[] declaredFields = castClass.getDeclaredFields();
        // TODO 加载类字段数据
    }

    private void getClassInfo() {

    }

    private Map<ReadExcelFeature, Object> defaultReadFeature() {
        Map<ReadExcelFeature, Object> defaultReadFeature = new HashMap<>();
        defaultReadFeature.put(IGNORE_EMPTY_SHEET_ERROR,IGNORE_EMPTY_SHEET_ERROR.getValue());
        defaultReadFeature.put(SORTED_READ_SHEET_DATA,SORTED_READ_SHEET_DATA.getValue());
        defaultReadFeature.put(INCLUDE_EMPTY_ROW,INCLUDE_EMPTY_ROW.getValue());
        defaultReadFeature.put(USE_MAP_DEBUG,USE_MAP_DEBUG.getValue());
        defaultReadFeature.put(DATA_BIND_PRECISE_LOCALIZATION,DATA_BIND_PRECISE_LOCALIZATION.getValue());
        defaultReadFeature.put(CAST_NUMBER_TO_DATE,CAST_NUMBER_TO_DATE.getValue());
        return defaultReadFeature;
    }

    public boolean getReadFeatureAsBoolean(ReadExcelFeature feature) {
        if (feature.getType() != ReadExcelFeature.Type.BOOLEAN){
            throw new IllegalArgumentException("读取特性不是一个布尔类型");
        }
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
                    if (getReadFeatureAsBoolean(SORTED_READ_SHEET_DATA)){
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

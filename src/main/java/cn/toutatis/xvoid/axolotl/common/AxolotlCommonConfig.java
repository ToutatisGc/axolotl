package cn.toutatis.xvoid.axolotl.common;

import cn.toutatis.xvoid.axolotl.Meta;
import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictKey;
import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictValue;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.xvoid.toolkit.clazz.ReflectToolkit;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.toolkit.validator.Validator;
import com.google.common.collect.HashBasedTable;
import lombok.Data;
import org.slf4j.Logger;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static java.lang.String.format;

@Data
public abstract class AxolotlCommonConfig {

    private Logger LOGGER = LoggerToolkit.getLogger(this.getClass());

    /**
     * sheet索引
     */
    protected int sheetIndex = 0;

    /**
     * 字典映射
     */
    protected HashBasedTable<Integer,String, Map<String,String>> dictionaryMapping = HashBasedTable.create();

    /**
     * Map映射指定键名
     */
    public static final String DICT_MAP_TYPE_POLICY_PREFIX = Meta.MODULE_NAME.toUpperCase()+"_DICT_MAPPING_POLICY_%s";
    public static final String DICT_MAP_TYPE_DEFAULT_PREFIX = Meta.MODULE_NAME.toUpperCase()+"_DICT_MAPPING_DEFAULT_%s";

    /**
     * 字典键值对名称指定
     */
    private String _dictKey = "key";
    private String _dictValue = "value";

    /**
     * 设置字典映射
     * @param field 字段
     * @param dict 字典
     */
    public void setDict(String field, List<?> dict) {
        this.setDict(getSheetIndex(), field, dict);
    }

    /**
     * 设置字典映射
     * @param field 字段
     * @param dict 字典
     */
    public void setDict(int sheetIndex,String field, List<?> dict) {
        if (Validator.strIsBlank(field)){
            throw new IllegalArgumentException("字段不能为空");
        }
        Map<String, Field> cache = new HashMap<>();
        if (dict != null && !dict.isEmpty()){
            Map<String,String> dictMap = new HashMap<>();
            Field keyField = null;
            Field valueField = null;
            for (Object item : dict) {
                if (item instanceof Map) {
                    Map<?, ?> itemMap = (Map<?, ?>)item;
                    if (itemMap.containsKey(_dictKey) && itemMap.containsKey(_dictValue)) {
                        dictMap.put(itemMap.get(_dictKey).toString(), itemMap.get(_dictValue).toString());
                    } else {
                        throw new AxolotlWriteException(LoggerHelper.format("请检查映射是否包含键值对名称[%s]:[%s]", _dictKey, _dictValue));
                    }
                } else {
                    if (cache.isEmpty()) {
                        List<Field> fields = ReflectToolkit.getAllFields(item.getClass(), true);
                        for (Field clazzField : fields) {
                            AxolotlDictKey keyAnno = clazzField.getAnnotation(AxolotlDictKey.class);
                            if (keyAnno != null) {
                                keyField = clazzField;
                                cache.put(get_dictKey(), clazzField);
                            }
                            AxolotlDictValue valueAnno = clazzField.getAnnotation(AxolotlDictValue.class);
                            if (valueAnno != null) {
                                valueField = clazzField;
                                cache.put(get_dictValue(), clazzField);
                            }
                        }
                        if (cache.size() != 2) {
                            throw new AxolotlWriteException("请检查实体字典映射是否包含注解@AxolotlDictKey和@AxolotlDictValue");
                        }
                    } else {
                        keyField = cache.get(get_dictKey());
                        valueField = cache.get(get_dictValue());
                    }
                    if (keyField == null || valueField == null) {
                        throw new AxolotlWriteException("请检查实体字典映射是否包含注解");
                    }
                    keyField.setAccessible(true);
                    valueField.setAccessible(true);
                    try {
                        Object code = keyField.get(item);
                        Object trans = valueField.get(item);
                        if (code != null && trans != null) {
                            dictMap.put(code.toString(), trans.toString());
                        } else {
                            throw new AxolotlWriteException("获取的字典为null，请检查实体值");
                        }
                    } catch (IllegalAccessException e) {
                        throw new RuntimeException(e);
                    }
                }
            }
            LoggerHelper.debug(LOGGER,format("字段[%s]字典映射数量[%s]", field,dictMap.size()));
            dictionaryMapping.put(sheetIndex,field,dictMap);
        }
    }

    /**
     * 设置字典映射
     * @param sheetIndex sheet索引
     * @param field 字段
     * @param dict 字典
     */
    public abstract void setDict(int sheetIndex,String field,Map<String,String> dict);

    /**
     * 获取字典映射
     * @param sheetIndex sheet索引
     * @param field 字段
     * @return 字典映射
     */
    public Map<String, String> getDict(int sheetIndex, String field) {
        Map<String, String> dict = dictionaryMapping.get(sheetIndex, field);
        if(dict == null){
            return new LinkedHashMap<>();
        }
        return dict;
    }

}

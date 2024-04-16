package cn.toutatis.xvoid.axolotl.excel.writer.support.base;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlDictKey;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlDictValue;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.exceptions.AxolotlException;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import com.google.common.collect.HashBasedTable;
import lombok.Data;
import lombok.SneakyThrows;
import org.slf4j.Logger;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static java.lang.String.format;

/**
 * 写入配置
 * @author Toutatis_Gc
 */
@Data
public class CommonWriteConfig {

    private Logger LOGGER = LoggerToolkit.getLogger(this.getClass());

    /**
     * 构造使用默认配置
     */
    public CommonWriteConfig() {
        this(true);
    }

    public CommonWriteConfig(boolean withDefaultConfig) {
        if (withDefaultConfig) {
            Map<ExcelWritePolicy, Object> defaultReadPolicies = new HashMap<>();
            for (ExcelWritePolicy policy : ExcelWritePolicy.values()) {
                if (policy.isDefaultPolicy()){
                    defaultReadPolicies.put(policy,policy.getValue());
                }
            }
            writePolicies.putAll(defaultReadPolicies);
        }
    }



    /**
     * sheet索引
     */
    private int sheetIndex = 0;

    /**
     * 写入策略
     */
    private Map<ExcelWritePolicy, Object> writePolicies = new HashMap<>();

    /**
     * 输出流
     */
    private OutputStream outputStream;

    /**
     * 字典映射
     */
    private HashBasedTable<Integer,String,Map<String,String>> dictionaryMapping = HashBasedTable.create();
    /**
     * 字典键值对名称指定
     */
    private String _dictKey = "key";
    private String _dictValue = "value";

    /**
     * 添加读取策略
     */
    public void setWritePolicy(ExcelWritePolicy policy,boolean value) {
        if (policy.getType() != ExcelWritePolicy.Type.BOOLEAN){
            throw new IllegalArgumentException("读取特性不是一个布尔类型");
        }
        writePolicies.put(policy,value);
    }

    /**
     * 获取一个布尔值类型的读取策略
     */
    public boolean getWritePolicyAsBoolean(ExcelWritePolicy policy) {
        if (policy.getType() != ExcelWritePolicy.Type.BOOLEAN){
            throw new IllegalArgumentException("读取特性不是一个布尔类型");
        }
        return writePolicies.containsKey(policy) && (boolean) writePolicies.get(policy);
    }

    /**
     * 设置字典映射
     * @param field 字段
     * @param dict 字典
     */
    public void setDict(int sheetIndex,String field, List<?> dict) {
        Map<String,Field> cache = new HashMap<>();
        if (dict != null && !dict.isEmpty()){
            Map<String,String> dictMap = new HashMap<>();
            Field keyField = null;
            Field valueField = null;
            for (Object item : dict) {
                if (item instanceof Map<?, ?> itemMap) {
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
     * @param field 字段
     * @param dict 字典
     */
    public void setDict(String field,Map<String,String> dict) {
        this.setDict(getSheetIndex(),field,dict);
    }

    /**
     * 设置字典映射
     * @param sheetIndex sheet索引
     * @param field 字段
     * @param dict 字典
     */
    public void setDict(int sheetIndex,String field,Map<String,String> dict) {
        dictionaryMapping.put(sheetIndex,field,dict);
    }

    /**
     * 获取字典映射
     * @param sheetIndex sheet索引
     * @param field 字段
     * @return 字典映射
     */
    public Map<String, String> getDict(int sheetIndex, String field) {
        return dictionaryMapping.get(sheetIndex,field);
    }
    /**
     * 关闭输出流
     */
    public void close() throws IOException {
        if (outputStream != null) {
            outputStream.close();
        }
    }

}

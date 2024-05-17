package cn.toutatis.xvoid.axolotl.excel.writer.support.base;

import cn.toutatis.xvoid.axolotl.Meta;
import cn.toutatis.xvoid.axolotl.common.AxolotlCommonConfig;
import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictKey;
import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictMapping;
import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictValue;
import cn.toutatis.xvoid.axolotl.common.annotations.DictMappingPolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.support.inverters.DataInverter;
import cn.toutatis.xvoid.axolotl.excel.writer.support.inverters.DefaultDataInverter;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import com.google.common.collect.HashBasedTable;
import lombok.Data;
import lombok.EqualsAndHashCode;
import org.slf4j.Logger;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy.SIMPLE_USE_DICT_CODE_TRANSFER;
import static java.lang.String.format;

/**
 * 写入配置
 * @author Toutatis_Gc
 */
@Data
@EqualsAndHashCode(callSuper = true)
public class CommonWriteConfig extends AxolotlCommonConfig {
    /**
     * 元数据类
     */
    protected Class<?> metaClass;

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
     * 空值填充字符
     * null值将被填充为空字符串，常用的字符串有"-","未填写","无"
     */
    private String blankValue = "";

    /**
     * 写入策略
     */
    private Map<ExcelWritePolicy, Object> writePolicies = new HashMap<>();

    /**
     * 输出流
     */
    private OutputStream outputStream;


    /**
     * 数据转换器
     */
    private DataInverter<?> dataInverter = new DefaultDataInverter();

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
    public void setDict(String field,Map<String,String> dict) {
        this.setDict(getSheetIndex(),field,dict);
    }

    /**
     * 关闭输出流
     */
    public void close() throws IOException {
        if (outputStream != null) {
            outputStream.close();
        }
    }

    @Override
    public void setDict(int sheetIndex, String field, Map<String, String> dict) {
        if (Validator.strIsBlank(field)){
            throw new IllegalArgumentException("字段不能为空");
        }
        if (dict!= null && !dict.isEmpty()){
            dictionaryMapping.put(sheetIndex,field,dict);
            boolean useDictCode = getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_USE_DICT_CODE_TRANSFER);
            if (!useDictCode){
                LoggerHelper.info(LOGGER,"字典映射策略未开启，已自动开启.");
                this.setWritePolicy(ExcelWritePolicy.SIMPLE_USE_DICT_CODE_TRANSFER,true);
            }
        }
    }

    /**
     * 检查实体类是否有字典注解，并自动开启字典策略
     */
    public void autoProcessEntity2OpenDictPolicy(){
        List<Field> allFields = ReflectToolkit.getAllFields(metaClass, true);
        Map<ExcelWritePolicy, Object> policyMap = getWritePolicies();
        if (!policyMap.containsKey(SIMPLE_USE_DICT_CODE_TRANSFER)){
            for (Field field : allFields) {
                if (field.getAnnotation(AxolotlDictMapping.class) != null){
                    LoggerHelper.info(LOGGER,"实体发现字典属性，字典映射策略未开启，已自动开启.");
                    this.setWritePolicy(SIMPLE_USE_DICT_CODE_TRANSFER,true);
                    break;
                }
            }
        }
    }
}

package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictMapping;
import cn.toutatis.xvoid.axolotl.common.annotations.DictMappingPolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.AxolotlSelectBox;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.AutoWriteContext;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.common.standard.StringPool;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;
import org.slf4j.Logger;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.*;

import static cn.toutatis.xvoid.axolotl.common.annotations.DictMappingPolicy.*;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

/**
 * 通用组件渲染器
 * @author Toutatis_Gc
 */
@Getter
public class ComponentRender {

    /**
     * 写入配置
     */
    @Setter
    protected AutoWriteConfig writeConfig;

    /**
     * 写入上下文
     */
    @Setter
    protected AutoWriteContext context;

    private final Logger LOGGER = LoggerToolkit.getLogger(this.getClass());

    /**
     * 渲染空值
     * @param fieldInfo 字段信息
     * @param cell 单元格
     */
    public void renderFieldColumnNullValue(AbstractStyleRender.FieldInfo fieldInfo, Cell cell){
        Object value = fieldInfo.getValue();
        if (value == null){
            cell.setCellValue(writeConfig.getBlankValue());
        }
    }

    /**
     * 渲染下拉框组件
     * @param fieldInfo 字段信息
     * @param cell 单元格
     */
    public void renderSelectDropListBox(AbstractStyleRender.FieldInfo fieldInfo, Cell cell){
        Object value = fieldInfo.getValue();
        AxolotlSelectBox<String> selectBox = ((AxolotlSelectBox<?>) value).convertPropertiesToString(writeConfig.getDataInverter());
        List<String> options = selectBox.getOptions();
        int switchSheetIndex = context.getSwitchSheetIndex();
        if(!options.isEmpty()){
            ExcelToolkit.createDropDownList(
                    context.getWorkbook().getSheetAt(switchSheetIndex),
                    selectBox.getOptions().toArray(new String[options.size()]),
                    fieldInfo.getRowIndex(),
                    fieldInfo.getRowIndex(),
                    fieldInfo.getColumnIndex(),
                    fieldInfo.getColumnIndex());
        }
        String selectBoxValue = selectBox.getValue();
        if(selectBoxValue == null){
            selectBoxValue = writeConfig.getBlankValue();
        }
        fieldInfo = new AbstractStyleRender.FieldInfo(
                fieldInfo.getDataInstance(),
                fieldInfo.getFieldName(),
                selectBoxValue,
                switchSheetIndex,
                fieldInfo.getColumnIndex(),
                fieldInfo.getRowIndex()
        );
        fieldInfo.setClazz(String.class);
        calculateColumns(fieldInfo);
        cell.setCellValue(selectBoxValue);
    }

    public void defaultRenderColumn(AbstractStyleRender.FieldInfo fieldInfo, Cell cell){
        Class<?> clazz = fieldInfo.getClazz();
        Object value = fieldInfo.getValue();
        // 组件渲染
        if(clazz == AxolotlSelectBox.class){
           this.renderSelectDropListBox(fieldInfo,cell);
        }else{
            calculateColumns(fieldInfo);
            value = writeConfig.getDataInverter().convert(value);
            //设置单元格值
            if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_USE_DICT_CODE_TRANSFER)){
                cell.setCellValue(convertDictCodeToName(fieldInfo, value.toString()));
            }else{
                cell.setCellValue(value.toString());
            }
        }
    }

    /**
     * 计算列合计
     * @param fieldInfo 字段信息
     */
    public void calculateColumns(AbstractStyleRender.FieldInfo fieldInfo){
        int columnIndex = fieldInfo.getColumnIndex();
        String value = fieldInfo.getValue().toString();
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING) && Validator.strIsNumber(value)){
            Map<Integer, BigDecimal> endingTotalMapping = context.getEndingTotalMapping().row(context.getSwitchSheetIndex());
            if (endingTotalMapping.containsKey(columnIndex)){
                BigDecimal newValue = endingTotalMapping.get(columnIndex).add(BigDecimal.valueOf(Double.parseDouble(value)));
                endingTotalMapping.put(columnIndex,newValue);
            }else{
                endingTotalMapping.put(columnIndex,BigDecimal.valueOf(Double.parseDouble(value)));
            }
        }
    }

    /**
     * 字典映射执行器
     */
    @Data
    private static class DictMappingExecutor{

        boolean isUsage = false;
        String defaultValue;
        boolean useManualConfigPriority = true;
        DictMappingPolicy mappingPolicy;
        private Map<String,String> convertMapping;

    }

    /**
     * 字典映射记录
     */
    private final HashMap<String,DictMappingExecutor> alreadyRecordDict = new HashMap<>();

    /**
     * 字典值转换编码到名称
     * @param fieldInfo 属性相关信息
     * @param value 单元格值
     * @return 字典值
     */
    @SuppressWarnings("unchecked")
    public String convertDictCodeToName(AbstractStyleRender.FieldInfo fieldInfo, String value){
        if (fieldInfo.getClazz() != String.class && fieldInfo.getClazz() != Integer.class){
            return value;
        }
        // 最终返回值
        String dictAdaptiveValue;
        int sheetIndex = fieldInfo.getSheetIndex();
        String fieldName = fieldInfo.getFieldName();
        Object dataInstance = fieldInfo.getDataInstance();
        Class<?> instanceClass = dataInstance.getClass();
        //字段键名称
        String alreadyKey = sheetIndex + StringPool.COLON + instanceClass.getSimpleName() + StringPool.COLON + fieldName;
        if (alreadyRecordDict.containsKey(alreadyKey)){
            debug(LOGGER,"字典值转换编码到名称: " + alreadyKey + " 已存在,直接返回.");
            DictMappingExecutor dictMappingExecutor = alreadyRecordDict.get(alreadyKey);
            if (dictMappingExecutor.isUsage){
                if (dictMappingExecutor.useManualConfigPriority){
                    Map<String, String> dict = writeConfig.getDict(sheetIndex, fieldName);
                    if (dict.containsKey(value)){
                        dictAdaptiveValue = dict.get(value);
                    }else{
                        Map<String, String> convertMapping = dictMappingExecutor.getConvertMapping();
                        if (convertMapping != null && convertMapping.containsKey(value)){
                            return convertMapping.get(value);
                        }else{
                            return adaptive(value,dictMappingExecutor.getMappingPolicy(),dictMappingExecutor.getDefaultValue());
                        }
                    }
                }else{
                    Map<String, String> convertMapping = dictMappingExecutor.getConvertMapping();
                    if (convertMapping != null && convertMapping.containsKey(value)){
                        return convertMapping.get(value);
                    }else{
                        return adaptive(value,dictMappingExecutor.getMappingPolicy(),dictMappingExecutor.getDefaultValue());
                    }
                }
            }else{
                return value;
            }
        }else{
            debug(LOGGER,"字典值转换编码到名称: " + alreadyKey + " 未存在,开始初始化.");
            DictMappingPolicy fieldDictMappingPolicy = null;
            String fieldDictMappingDefaultValue = null;
            DictMappingExecutor dictMappingExecutor = new DictMappingExecutor();
            if (dataInstance instanceof Map<?, ?>){
                if(fieldName != null){
                    Map<String, Object> instanceMap = (Map<String, Object>) dataInstance;
                    //获取字典策略
                    String policyKey = CommonWriteConfig.DICT_MAP_TYPE_POLICY_PREFIX.formatted(fieldName);
                    if (instanceMap.containsKey(policyKey)){
                        Object fieldPolicy = instanceMap.get(policyKey);
                        if (fieldPolicy instanceof DictMappingPolicy){
                            fieldDictMappingPolicy = (DictMappingPolicy) fieldPolicy;
                        }else{
                            try {
                                fieldDictMappingPolicy = DictMappingPolicy.valueOf(fieldPolicy.toString());
                            }catch (Exception ex){
                                ex.printStackTrace();
                                if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_EXCEPTION_RETURN_RESULT)){
                                    error(LOGGER, format("枚举转换异常，未知的枚举[%s]", fieldPolicy.toString()));
                                }else{
                                    throw ex;
                                }
                            }
                        }
                    }
                    String defaultValueKey = CommonWriteConfig.DICT_MAP_TYPE_DEFAULT_PREFIX.formatted(fieldName);
                    if (instanceMap.containsKey(defaultValueKey)){
                        Object defaultValueObject =  instanceMap.get(defaultValueKey);
                        if (defaultValueObject != null){
                            fieldDictMappingDefaultValue = defaultValueObject.toString();
                        }
                    }
                }else {
                    warn(LOGGER,"Map写入字段名称为空,无法获取字段信息。");
                }
                // POJO类型
            }else{
                Field field = ReflectToolkit.recursionGetField(dataInstance.getClass(),fieldName);
                // 字段不存在查找getter方法
                AxolotlDictMapping axolotlDictMapping = null;
                if (field == null){
                    String fieldGetterMethodName = ReflectToolkit.getFieldGetterMethodName(fieldName);
                    try {
                        Method getterMethod = fieldInfo.getDataInstance().getClass().getDeclaredMethod(fieldGetterMethodName);
                        axolotlDictMapping = getterMethod.getAnnotation(AxolotlDictMapping.class);
                    } catch (NoSuchMethodException e) {
                        throw new RuntimeException(e);
                    }
                }
                if (axolotlDictMapping == null){
                    if (field != null){
                        axolotlDictMapping = field.getAnnotation(AxolotlDictMapping.class);
                    }
                }
                if (axolotlDictMapping != null){
                    // 分配实体策略
                    dictMappingExecutor.setUsage(axolotlDictMapping.isUsage());
                    fieldDictMappingPolicy = axolotlDictMapping.mappingPolicy();
                    if (axolotlDictMapping.defaultValue().length >= 1) {
                        fieldDictMappingDefaultValue = axolotlDictMapping.defaultValue()[0];
                    }
                    // 将字典值加入到config配置
                    String[] staticDictArray = axolotlDictMapping.staticDict();
                    if (staticDictArray.length > 0){
                        if (staticDictArray.length % 2 == 0){
                            HashMap<String, String> staticMap = new HashMap<>(staticDictArray.length / 2);
                            for (int i = 0; i < staticDictArray.length; i += 2) {
                                staticMap.put(staticDictArray[i], staticDictArray[i + 1]);
                            }
                            dictMappingExecutor.setConvertMapping(staticMap);
                        }else{
                            throw new IllegalArgumentException("静态字典数组长度必须为偶数");
                        }
                    }
                }
            }
            // 统一策略阶段空值使用默认策略
            if (fieldDictMappingPolicy == null){
                debug(LOGGER, "未获取到字典策略，字段[%s]使用默认策略[%s]",fieldName, KEEP_ORIGIN);
                fieldDictMappingPolicy = KEEP_ORIGIN;
                dictMappingExecutor.isUsage = true;
            }else{
                debug(LOGGER, "字段[%s]使用字典策略[%s]",fieldName, fieldDictMappingPolicy);
            }
            if (fieldDictMappingDefaultValue == null){
                debug(LOGGER, "未获取到字典默认值，字段[%s]使用配置默认值[%s]",fieldName, writeConfig.getBlankValue());
                dictMappingExecutor.setDefaultValue(writeConfig.getBlankValue());
            }else{
                dictMappingExecutor.setDefaultValue(fieldDictMappingDefaultValue);
            }
            dictMappingExecutor.setMappingPolicy(fieldDictMappingPolicy);
            alreadyRecordDict.put(alreadyKey, dictMappingExecutor);
            return convertDictCodeToName(fieldInfo,value);
        }
        return dictAdaptiveValue;
    }

    private String adaptive(String value, DictMappingPolicy fieldDictMappingPolicy, String fieldDictMappingDefaultValue){
        return switch (fieldDictMappingPolicy) {
            case KEEP_ORIGIN -> value;
            case USE_DEFAULT -> fieldDictMappingDefaultValue;
            case NULL_VALUE -> null;
        };
    }

}

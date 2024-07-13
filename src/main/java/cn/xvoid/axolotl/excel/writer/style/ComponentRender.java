package cn.xvoid.axolotl.excel.writer.style;

import cn.xvoid.axolotl.common.AxolotlCommonConfig;
import cn.xvoid.axolotl.common.annotations.AxolotlDictMapping;
import cn.xvoid.axolotl.common.annotations.AxolotlDictOverTurn;
import cn.xvoid.axolotl.common.annotations.DictMappingPolicy;
import cn.xvoid.axolotl.excel.writer.components.widgets.AxolotlSelectBox;
import cn.xvoid.axolotl.excel.writer.support.base.AutoWriteContext;
import cn.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
import cn.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.xvoid.axolotl.excel.writer.support.base.WriteContext;
import cn.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.xvoid.common.standard.StringPool;
import cn.xvoid.toolkit.clazz.ReflectToolkit;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.toolkit.validator.Validator;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.util.*;

import static cn.xvoid.axolotl.toolkit.LoggerHelper.*;

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
    protected AxolotlCommonConfig config;

    /**
     * 写入上下文
     */
    @Setter
    protected WriteContext context;

    /**
     * 日志
     */
    private final Logger LOGGER = LoggerToolkit.getLogger(this.getClass());

    @Setter
    private boolean isReader = false;

    /**
     * 渲染空值
     * @param fieldInfo 字段信息
     * @param cell 单元格
     */
    public void renderFieldColumnNullValue(AbstractStyleRender.FieldInfo fieldInfo, Cell cell){
        Object value = fieldInfo.getValue();
        if (value == null){
            cell.setCellValue(getCommonWriteConfig().getBlankValue());
        }
    }

    /**
     * 渲染下拉框组件
     * @param fieldInfo 字段信息
     * @param cell 单元格
     */
    public void renderSelectDropListBox(AbstractStyleRender.FieldInfo fieldInfo, Cell cell){
        Object value = fieldInfo.getValue();
        AxolotlSelectBox<String> selectBox = ((AxolotlSelectBox<?>) value).convertPropertiesToString(getCommonWriteConfig().getDataInverter());
        List<String> options = selectBox.getOptions();
        int switchSheetIndex = context.getSwitchSheetIndex();
        Sheet sheet;
        if(context instanceof AutoWriteContext){
            AutoWriteContext autoWriteContext = (AutoWriteContext) context;
            sheet = autoWriteContext.getWorkbook().getSheetAt(switchSheetIndex);
        }else{
            throw new IllegalArgumentException("该功能为自动写入功能,请使用AutoWriteContext");
        }
        if(!options.isEmpty()){
            ExcelToolkit.createDropDownList(
                    sheet,
                    selectBox.getOptions().toArray(new String[options.size()]),
                    fieldInfo.getRowIndex(),
                    fieldInfo.getRowIndex(),
                    fieldInfo.getColumnIndex(),
                    fieldInfo.getColumnIndex());
        }
        String selectBoxValue = selectBox.getValue();
        if(selectBoxValue == null){
            selectBoxValue = getCommonWriteConfig().getBlankValue();
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
            value = getCommonWriteConfig().getDataInverter().convert(value);
            //设置单元格值
            if (getCommonWriteConfig().getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_USE_DICT_CODE_TRANSFER)){
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
        if (getCommonWriteConfig().getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING) && Validator.strIsNumber(value)){
            Map<Integer, BigDecimal> endingTotalMapping;
            if (context instanceof AutoWriteContext){
                AutoWriteContext autoWriteContext = (AutoWriteContext) context;
                endingTotalMapping = autoWriteContext.getEndingTotalMapping().row(context.getSwitchSheetIndex());
            }else {
                throw new IllegalArgumentException("该功能为自动写入功能,请使用AutoWriteContext");
            }
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
    public String convertDictCodeToName(AbstractStyleRender.FieldInfo fieldInfo, String value){
        return convertDictCodeToName(fieldInfo.getSheetIndex(),fieldInfo.getClazz(),fieldInfo.getFieldName(),fieldInfo.getDataInstance(),value);
    }

    /**
     * 字典值转换编码到名称
     * @param sheetIndex sheet索引
     * @param fieldClass 字段类型
     * @param fieldName 字段名称
     * @param dataInstance 数据实例
     * @param value 单元格值
     * @return 字典值
     */
    @SuppressWarnings("unchecked")
    public String convertDictCodeToName(int sheetIndex,Class<?> fieldClass,String fieldName,Object dataInstance, String value){
        if (fieldClass != String.class
                && fieldClass != Integer.class && fieldClass != int.class
                && fieldClass != Boolean.class && fieldClass != boolean.class
        ){
            return value;
        }
        // 最终返回值
        String dictAdaptiveValue;
        Class<?> instanceClass = dataInstance.getClass();
        //字段键名称
        String alreadyKey = sheetIndex + StringPool.COLON + instanceClass.getSimpleName() + StringPool.COLON + fieldName;
        if (alreadyRecordDict.containsKey(alreadyKey)){
            LoggerHelper.debug(LOGGER,"字典值转换编码到名称: " + alreadyKey + " 已存在,直接返回.");
            DictMappingExecutor dictMappingExecutor = alreadyRecordDict.get(alreadyKey);
            if (dictMappingExecutor.isUsage){
                if (dictMappingExecutor.useManualConfigPriority){
                    Map<String, String> dict = getCommonWriteConfig().getDict(sheetIndex, fieldName);
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
                    String policyKey = String.format(CommonWriteConfig.DICT_MAP_TYPE_POLICY_PREFIX, fieldName);
                    if (instanceMap.containsKey(policyKey)){
                        Object fieldPolicy = instanceMap.get(policyKey);
                        if (fieldPolicy instanceof DictMappingPolicy){
                            fieldDictMappingPolicy = (DictMappingPolicy) fieldPolicy;
                        }else{
                            try {
                                fieldDictMappingPolicy = DictMappingPolicy.valueOf(fieldPolicy.toString());
                            }catch (Exception ex){
                                ex.printStackTrace();
                                if (getCommonWriteConfig().getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_EXCEPTION_RETURN_RESULT)){
                                    error(LOGGER, format("枚举转换异常，未知的枚举[%s]", fieldPolicy.toString()));
                                }else{
                                    throw ex;
                                }
                            }
                        }
                    }
                    String defaultValueKey = String.format(CommonWriteConfig.DICT_MAP_TYPE_DEFAULT_PREFIX, fieldName);
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
                AxolotlDictOverTurn axolotlDictOverTurn = null;
                if (field == null){
                    String fieldGetterMethodName = ReflectToolkit.getFieldGetterMethodName(fieldName);
                    try {
                        Method getterMethod = dataInstance.getClass().getDeclaredMethod(fieldGetterMethodName);
                        axolotlDictMapping = getterMethod.getAnnotation(AxolotlDictMapping.class);
                        axolotlDictOverTurn = getterMethod.getAnnotation(AxolotlDictOverTurn.class);
                    } catch (NoSuchMethodException e) {
                        throw new RuntimeException(e);
                    }
                }
                if (axolotlDictMapping == null){
                    if (field != null){
                        axolotlDictMapping = field.getAnnotation(AxolotlDictMapping.class);
                        axolotlDictOverTurn = field.getAnnotation(AxolotlDictOverTurn.class);
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
                            boolean overTurn = false;
                            if (axolotlDictMapping.autoOverTurn()){
                                overTurn = true;
                            }
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
                debug(LOGGER, "未获取到字典策略，字段[%s]使用默认策略[%s]",fieldName, DictMappingPolicy.KEEP_ORIGIN);
                fieldDictMappingPolicy = DictMappingPolicy.KEEP_ORIGIN;
                dictMappingExecutor.isUsage = true;
            }else{
                debug(LOGGER, "字段[%s]使用字典策略[%s]",fieldName, fieldDictMappingPolicy);
            }
            if (fieldDictMappingDefaultValue == null){
                debug(LOGGER, "未获取到字典默认值，字段[%s]使用配置默认值[%s]",fieldName, getCommonWriteConfig().getBlankValue());
                dictMappingExecutor.setDefaultValue(getCommonWriteConfig().getBlankValue());
            }else{
                dictMappingExecutor.setDefaultValue(fieldDictMappingDefaultValue);
            }
            dictMappingExecutor.setMappingPolicy(fieldDictMappingPolicy);
            alreadyRecordDict.put(alreadyKey, dictMappingExecutor);
            return convertDictCodeToName(sheetIndex,fieldClass,fieldName,dataInstance,value);
        }
        return dictAdaptiveValue;
    }

    private String adaptive(String value, DictMappingPolicy fieldDictMappingPolicy, String fieldDictMappingDefaultValue){
        String result;
        switch (fieldDictMappingPolicy) {
            case KEEP_ORIGIN:
                result = value;
                break;
            case USE_DEFAULT:
                result = fieldDictMappingDefaultValue;
                break;
            case NULL_VALUE:
                result = null;
                break;
            default:
                return null;
        }
// 使用结果
        return result;
    }

    private CommonWriteConfig getCommonWriteConfig(){
        return (CommonWriteConfig) config;
    }

    private boolean isWriteConfig(){
        return config instanceof CommonWriteConfig;
    }

}

package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictMapping;
import cn.toutatis.xvoid.axolotl.common.annotations.DictMappingPolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.AxolotlSelectBox;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.AutoWriteContext;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.FieldToolkit;
import cn.toutatis.xvoid.common.standard.StringPool;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import com.google.common.collect.HashBasedTable;
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
        if(!options.isEmpty()){
            ExcelToolkit.createDropDownList(
                    context.getWorkbook().getSheetAt(context.getSwitchSheetIndex()),
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
            // TODO GET方法,用ReflectToolkit,同步你的可配置主题相同代码
            // TODO 字典转换
            value = writeConfig.getDataInverter().convert(value);
            //设置单元格值
            cell.setCellValue(convertDictCodeToName(fieldInfo, value.toString()));
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

    private final HashMap<String,DictMappingPolicy> alreadyRecordDict = new HashMap<>();

//    private final HashBasedTable<Integer>
    /**
     * 字典值转换编码到名称
     * map类型只支持设置全局字典
     * @param fieldInfo 属性相关信息
     * @param value 单元格值
     * @return 字典值
     */
    @SuppressWarnings("unchecked")
    public String convertDictCodeToName(AbstractStyleRender.FieldInfo fieldInfo, String value){
        int sheetIndex = context.getSwitchSheetIndex();
        String fieldName = fieldInfo.getFieldName();
        Object dataInstance = fieldInfo.getDataInstance();
        Class<?> instanceClass = dataInstance.getClass();
        boolean isMap = dataInstance instanceof Map<?, ?>;
        String alreadyKey = sheetIndex + StringPool.COLON + instanceClass + StringPool.COLON + fieldName;
        Map<String, String> dict = writeConfig.getDict(sheetIndex, fieldName);
        if (dict.containsKey(value)){
            return dict.get(value);
        }else{
//            String alreadyKeyWithAlias =
            DictMappingPolicy fieldDictMappingPolicy = null;
            String dictAdaptiveValue = value;
            String fieldDictMappingDefaultValue = null;
            if (alreadyRecordDict.containsKey(alreadyKey)){
                fieldDictMappingPolicy = alreadyRecordDict.get(alreadyKey);
            }else{
                // Map类型从key中取值
                if (isMap){
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
                    Field field = FieldToolkit.recursionGetField(dataInstance.getClass(),fieldName);
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
                    if (axolotlDictMapping != null && axolotlDictMapping.isUsage()){
                        // 分配实体策略
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
                                writeConfig.setDict(fieldName, staticMap);
                            }else{
                                throw new IllegalArgumentException("静态字典数组长度必须为偶数");
                            }
                        }
                    }
                }
            }
            // 统一策略阶段空值使用默认策略
            if (fieldDictMappingPolicy == null){
                debug(LOGGER, "未获取到字典策略，字段[%s]使用默认策略[%s]",fieldName, KEEP_ORIGIN);
                fieldDictMappingPolicy = KEEP_ORIGIN;
            }else{
                debug(LOGGER, "字段[%s]使用字典策略[%s]",fieldName, fieldDictMappingPolicy);
            }
            dictAdaptiveValue = switch (fieldDictMappingPolicy) {
                case KEEP_ORIGIN -> dictAdaptiveValue;
                case USE_DEFAULT -> fieldDictMappingDefaultValue;
                case NULL_VALUE -> null;
            };
            return dictAdaptiveValue;
        }
    }

    // if(!(dataInstance instanceof Map<?,?>)){
    //            Field field = FieldToolkit.recursionGetField(dataInstance.getClass(), fieldInfo.getFieldName());
    //            if(field != null){
    //                AxolotlDictMapping dictMappingInfo = field.getAnnotation(AxolotlDictMapping.class);
    //                if (dictMappingInfo != null){
    //                    //带有 AxolotlDictMapping 注解的处理
    //                    //是否进行字典值处理
    //                    if(dictMappingInfo.isUsage()){
    //                        int[] sheetIndexs = dictMappingInfo.effectSheetIndex();
    //                        boolean isUsage = false;
    //                        if(sheetIndexs.length != 0){
    //                            for (int index : sheetIndexs) {
    //                                if(index == sheetIndex){
    //                                    isUsage = true;
    //                                    break;
    //                                }
    //                            }
    //                        }else{
    //                            isUsage = true;
    //                        }
    //                        if(isUsage){
    //                            //映射字段名
    //                            String mappingFieldName;
    //                            if(StringUtils.isNotEmpty(dictMappingInfo.value())){
    //                                mappingFieldName = dictMappingInfo.value();
    //                            }else{
    //                                mappingFieldName = fieldInfo.getFieldName();
    //                            }
    //                            Map<String, String> dictMapping = writeConfig.getDict(sheetIndex, mappingFieldName);
    //                            //静态字典
    //                            if(dictMappingInfo.staticDict().length > 0 && dictMappingInfo.staticDict().length%2 == 0){
    //                                //使用静态字典处理
    //                                dictMapping = new LinkedHashMap<>();
    //                                for (int i = 0; i < dictMappingInfo.staticDict().length; i++) {
    //                                    if (i % 2 != 0) {
    //                                        dictMapping.put(dictMappingInfo.staticDict()[i-1],dictMappingInfo.staticDict()[i]);
    //                                    }
    //                                }
    //                            }
    //                            if(!dictMapping.isEmpty()){
    //                                String dictName = dictMapping.get(value);
    //                                if(dictName == null){
    //                                    AxolotlDictMappingPolicy dictMappingPolicyAnno = field.getAnnotation(AxolotlDictMappingPolicy.class);
    //                                    if(dictMappingPolicyAnno != null){
    //                                        if(dictMappingPolicyAnno.mappingPolicy().equals(KEEP_ORIGIN)){
    //                                            //保留字段原值
    //                                            //  value = value.toString();
    //                                        }else if(dictMappingPolicyAnno.mappingPolicy().equals(USE_DEFAULT)){
    //                                            //使用字段默认值
    //                                            value = dictMappingPolicyAnno.defaultValue();
    //                                        }else if(dictMappingPolicyAnno.mappingPolicy().equals(NULL_VALUE)){
    //                                            //设置为空
    //                                            value = null;
    //                                        }
    //                                    }else{
    //                                        Map<DictMappingPolicy, String> dictMappingPolicy = writeConfig.getDictMappingPolicy(sheetIndex);
    //                                        if(dictMappingPolicy.containsKey(KEEP_ORIGIN)){
    //                                            //保留字段原值
    //                                            //  value = value.toString();
    //                                        }else if(dictMappingPolicy.containsKey(USE_DEFAULT)){
    //                                            //使用字段默认值
    //                                            value = dictMappingPolicy.get(USE_DEFAULT);
    //                                        }else if(dictMappingPolicy.containsKey(NULL_VALUE)){
    //                                            //设置为空
    //                                            value = null;
    //                                        }
    //                                    }
    //                                }else{
    //                                    value = dictName;
    //                                }
    //                            }
    //                        }
    //                    }
    //                }else{
    //                    //没有 AxolotlDictMapping 注解的处理
    //                    Map<String, String> dictMapping = writeConfig.getDict(sheetIndex, fieldInfo.getFieldName());
    //                    if(!dictMapping.isEmpty()){
    //                        String dictName = dictMapping.get(value);
    //                        if(dictName == null){
    //                            AxolotlDictMappingPolicy dictMappingPolicyAnno = field.getAnnotation(AxolotlDictMappingPolicy.class);
    //                            if(dictMappingPolicyAnno != null){
    //                                if(dictMappingPolicyAnno.mappingPolicy().equals(KEEP_ORIGIN)){
    //                                    //保留字段原值
    //                                    //  value = value.toString();
    //                                }else if(dictMappingPolicyAnno.mappingPolicy().equals(USE_DEFAULT)){
    //                                    //使用字段默认值
    //                                    value = dictMappingPolicyAnno.defaultValue();
    //                                }else if(dictMappingPolicyAnno.mappingPolicy().equals(NULL_VALUE)){
    //                                    //设置为空
    //                                    value = null;
    //                                }
    //                            }else{
    //                                Map<DictMappingPolicy, String> dictMappingPolicy = writeConfig.getDictMappingPolicy(sheetIndex);
    //                                if(dictMappingPolicy.containsKey(KEEP_ORIGIN)){
    //                                    //保留字段原值
    //                                    //  value = value.toString();
    //                                }else if(dictMappingPolicy.containsKey(USE_DEFAULT)){
    //                                    //使用字段默认值
    //                                    value = dictMappingPolicy.get(USE_DEFAULT);
    //                                }else if(dictMappingPolicy.containsKey(NULL_VALUE)){
    //                                    //设置为空
    //                                    value = null;
    //                                }
    //                            }
    //                        }else{
    //                            value = dictName;
    //                        }
    //                    }
    //                }
    //            }else{
    //                //getter方法 使用全局字典值映射策略
    //                Map<String, String> dictMapping = writeConfig.getDict(sheetIndex, fieldInfo.getFieldName());
    //                if(!dictMapping.isEmpty()){
    //                    String dictName = dictMapping.get(value);
    //                    if(dictName == null){
    //                        Map<DictMappingPolicy, String> dictMappingPolicy = writeConfig.getDictMappingPolicy(sheetIndex);
    //                        if(dictMappingPolicy.containsKey(KEEP_ORIGIN)){
    //                            //保留字段原值
    //                            //  value = value.toString();
    //                        }else if(dictMappingPolicy.containsKey(USE_DEFAULT)){
    //                            //使用字段默认值
    //                            value = dictMappingPolicy.get(USE_DEFAULT);
    //                        }else if(dictMappingPolicy.containsKey(NULL_VALUE)){
    //                            //设置为空
    //                            value = null;
    //                        }
    //                    }else{
    //                        value = dictName;
    //                    }
    //                }
    //            }
    //        }
    //        return dictAdaptiveValue;


}

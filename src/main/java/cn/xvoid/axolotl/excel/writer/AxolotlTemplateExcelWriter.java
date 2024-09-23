package cn.xvoid.axolotl.excel.writer;

import cn.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.xvoid.axolotl.excel.writer.components.annotations.AxolotlWriteIgnore;
import cn.xvoid.axolotl.excel.writer.components.annotations.AxolotlWriterGetter;
import cn.xvoid.axolotl.excel.writer.constant.TemplatePlaceholderPattern;
import cn.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.xvoid.axolotl.excel.writer.style.ComponentRender;
import cn.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.xvoid.axolotl.excel.writer.support.base.*;
import cn.xvoid.axolotl.excel.writer.support.inverters.DataInverter;
import cn.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.xvoid.axolotl.toolkit.LoggerHelper;
import cn.xvoid.axolotl.toolkit.tika.TikaShell;
import cn.xvoid.common.standard.StringPool;
import cn.xvoid.toolkit.clazz.ReflectToolkit;
import cn.xvoid.toolkit.constant.Time;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.toolkit.validator.Validator;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.MapDifference;
import com.google.common.collect.Maps;
import lombok.Getter;
import lombok.Setter;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;

import java.io.File;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 模板文档文件写入器
 * @author Toutatis_Gc
 */
public class AxolotlTemplateExcelWriter extends AxolotlAbstractExcelWriter {

    /**
     * 日志工具
     * 日志记录器
     */
    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlTemplateExcelWriter.class);

    /**
     * 写入配置
     */
    private final TemplateWriteConfig config;

    /**
     * 写入上下文
     */
    private final TemplateWriteContext context;

    /**
     * 组件渲染器
     */
    @Getter @Setter
    private ComponentRender componentRender = new ComponentRender() {};

    /**
     * 主构造函数
     *
     * @param templateWriteConfig 写入配置
     */
    public AxolotlTemplateExcelWriter(TemplateWriteConfig templateWriteConfig) {
        super.LOGGER = LOGGER;
        this.config = templateWriteConfig;
        this.checkWriteConfig(this.config);
        TemplateWriteContext templateWriteContext = new TemplateWriteContext();
        super.writeContext = templateWriteContext;
        this.context = templateWriteContext;
        this.context.setSwitchSheetIndex(config.getSheetIndex());
        componentRender.setContext(context);
        componentRender.setConfig(config);
    }

    /**
     * 构造函数
     * 可以写入一个模板文件
     *
     * @param templateFile 模板文件
     * @param config 写入配置
     */
    public AxolotlTemplateExcelWriter(File templateFile, TemplateWriteConfig config) {
        this(config);
        TikaShell.preCheckFileNormalThrowException(templateFile);
        this.workbook = this.initWorkbook(templateFile);
    }

    /**
     * 写入Excel数据
     * @param fixMapping 固定值数据
     * @param dataList 循环引用数据
     * @return 写入结果
     * @throws AxolotlWriteException 写入异常
     */
    public AxolotlWriteResult write(Map<String, ?> fixMapping, List<?> dataList) {
        LoggerHelper.info(LOGGER, context.getCurrentWrittenBatchAndIncrement(context.getSwitchSheetIndex()));
        Sheet sheet;
        // 判断是否是模板写入
        AxolotlWriteResult axolotlWriteResult = new AxolotlWriteResult();
        if (context.isTemplateWrite()){
            int switchSheetIndex = context.getSwitchSheetIndex();
            sheet = getWorkbookSheet(switchSheetIndex);
            // 只有第一次写入时解析模板占位符
            if (context.isFirstBatch(switchSheetIndex)){
                // 解析模板占位符到上下文
                this.resolveTemplate(sheet,false);
            }
//            // 写入Map映射
            this.writeSingleData(sheet, fixMapping, context.getSingleReferenceData(),false);
//            // 写入循环数据
            if (Validator.objNotNull(dataList)){
                config.setMetaClass(dataList.get(0).getClass());
                config.autoProcessEntity2OpenDictPolicy();
            }
            this.writeCircleData(sheet, dataList);
            axolotlWriteResult.setWrite(true);
            axolotlWriteResult.setMessage("写入完成");
        }else{
            String message = "非模板写入请使用AxolotlAutoExcelWriter.write()方法";
            if(config.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_EXCEPTION_RETURN_RESULT)){
                axolotlWriteResult.setMessage(message);
                return axolotlWriteResult;
            }
            throw new AxolotlWriteException(message);
        }
        return axolotlWriteResult;
    }

    /**
     * 写入单次引用数据
     *
     * @param sheet 工作表
     * @param singleMap 单次引用数据
     * @param referenceData 数据源
     * @param gatherUnusedStage 是否收集未使用的数据
     */
    private void writeSingleData(Sheet sheet, Map<String,?> singleMap, HashBasedTable<Integer, String, CellAddress> referenceData, boolean gatherUnusedStage){
        // 注入通用信息
        HashMap<String, Object> dataMapping = (singleMap != null ? new HashMap<>(singleMap) : new HashMap<>());
        this.injectCommonConstInfo(dataMapping,gatherUnusedStage);
        int sheetIndex = getSheetIndex(sheet);
        Map<String, CellAddress> addressMapping = referenceData.row(sheetIndex);
        // 记录已使用的引用数据
        Map<String, Boolean> alreadyUsedReferenceData = context.getAlreadyUsedReferenceData().row(sheetIndex);
        for (String singleKey : addressMapping.keySet()) {
            // 如果地址引用包含该关键字，则写入数据
            CellAddress cellAddress = addressMapping.get(singleKey);
            String placeholder = cellAddress.getPlaceholder();
            if(dataMapping.containsKey(singleKey)){
                // 已经写入过则跳过写入
                if(alreadyUsedReferenceData.containsKey(placeholder)){
                    LoggerHelper.debug(LOGGER, LoggerHelper.format("已跳过使用的占位符[%s]",placeholder));
                    continue;
                }
                // 写入单元格值
                Cell cell = sheet.getRow(cellAddress.getRowPosition()).getCell(cellAddress.getColumnPosition());
                String stringCellValue = cell.getStringCellValue();
                Object info = dataMapping.get(singleKey);
                if(info != null){
                    LoggerHelper.debug(LOGGER, LoggerHelper.format("设置模板占位符[%s]值[%s]",placeholder,info));
                    String replaceString = stringCellValue.replace(placeholder, info.toString());
                    cell.setCellValue(replaceString);
                }else{
                    String defaultValue = cellAddress.getDefaultValue();
                    String newCellValue = null;
                    boolean isDefaultValue = false;
                    if(defaultValue != null){
                        isDefaultValue = true;
                        newCellValue = stringCellValue.replace(placeholder, defaultValue);
                    }else{
                        if(!stringCellValue.equals(placeholder)){
                            newCellValue = stringCellValue.replace(placeholder, "");
                        }
                    }
                    if (newCellValue == null){
                        LoggerHelper.debug(LOGGER, LoggerHelper.format("%s设置模板占位符[%s]为空值",gatherUnusedStage ? "[收尾阶段]":config.getBlankValue(),placeholder));
                        cell.setBlank();
                    }else{
                        LoggerHelper.debug(LOGGER, LoggerHelper.format("%s设置模板占位符[%s]为[%s]值",gatherUnusedStage ? "[收尾阶段]":"", placeholder,isDefaultValue ? "默认":"空"));
                        cell.setCellValue(newCellValue);
                    }
                }
                cellAddress.setWrittenRow(cell.getRowIndex());
                alreadyUsedReferenceData.put(placeholder,true);
            }else{
                LoggerHelper.debug(LOGGER, LoggerHelper.format("未使用模板占位符[%s]",placeholder));
            }
        }
    }

    /**
     * 未使用的单次占位符填充默认值
     */
    private void gatherUnusedSingleReferenceDataAndFillDefault() {
        int sheetIndex = context.getSwitchSheetIndex();
        Sheet sheet = this.getWorkbookSheet(context.getSwitchSheetIndex());
        Map<String, CellAddress> singleReferenceMapping =  context.getSingleReferenceData().row(sheetIndex);
        HashMap<String, CellAddress> tmp = singleReferenceMapping.entrySet()
                .stream()
                .collect(Collectors.toMap(
                        entry -> entry.getValue().getPlaceholder(), Map.Entry::getValue,
                        (a, b) -> b, () -> new HashMap<>(singleReferenceMapping.size()))
                );
        HashMap<String, Object> unusedMap = gatherUnusedField(sheetIndex, tmp);
        this.writeSingleData(sheet,unusedMap, context.getSingleReferenceData(),true);
    }

    /**
     * 未使用的列表占位符填充默认值
     */
    private void gatherUnusedCircleReferenceDataAndFillDefault() {
        if(config.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_PLACEHOLDER_FILL_DEFAULT)){
            int sheetIndex = context.getSwitchSheetIndex();
            Sheet sheet = this.getWorkbookSheet(sheetIndex);
            Map<String, CellAddress> circleReferenceData =  context.getCircleReferenceData().row(sheetIndex);
            HashMap<String, Object> map = gatherUnusedField(sheetIndex, circleReferenceData);
            this.writeSingleData(sheet,map, context.getCircleReferenceData(),true);
        }
    }

    /**
     * 设置计算数据
     */
    private void setCalculateData(int sheetIndex) {
        HashBasedTable<Integer, String, CellAddress> calculateReferenceData = context.getCalculateReferenceData();
        Map<String, CellAddress> calculateData = calculateReferenceData.row(sheetIndex);
        Map<String, BigDecimal> extract = new HashMap<>();
        DataInverter<?> dataInverter = config.getDataInverter();
        boolean calculateIgnoreDecimal = config.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_CALCULATE_INTEGER_IGNORE_DECIMAL);
        for (String s : calculateData.keySet()) {
            BigDecimal calculatedValue = calculateData.get(s).getCalculatedValue();
            boolean isInteger = false;
            if (calculateIgnoreDecimal){
                isInteger = calculatedValue.remainder(BigDecimal.ONE).compareTo(BigDecimal.ZERO) == 0;
            }
            extract.put(s,calculatedValue.setScale(isInteger ? 0 : AxolotlDefaultReaderConfig.XVOID_DEFAULT_DECIMAL_SCALE, RoundingMode.HALF_UP));
        }
        this.writeSingleData(getWorkbookSheet(sheetIndex),extract,calculateReferenceData,true);
    }

    /**
     * 采集未使用的占位符
     *
     * @param sheetIndex 工作表索引
     * @param referenceMapping 数据源
     * @return 未使用的占位符
     */
    private HashMap<String, Object> gatherUnusedField(int sheetIndex, Map<String, CellAddress> referenceMapping) {
        Map<String, Boolean> alreadyUsedDataMapping =  context.getAlreadyUsedReferenceData().row(sheetIndex);
        MapDifference<String, Object> difference = Maps.difference(referenceMapping, alreadyUsedDataMapping);
        Map<String, Object> onlyOnLeft = difference.entriesOnlyOnLeft();
        HashMap<String, Object> unusedMap = new HashMap<>();
        for (Map.Entry<String, Object> entry : onlyOnLeft.entrySet()) {
            CellAddress cellAddress = (CellAddress) entry.getValue();
            Object nonAddressValue;
            if (config.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_PLACEHOLDER_FILL_DEFAULT)){
                if(cellAddress.getDefaultValue() != null){
                    nonAddressValue = cellAddress.getDefaultValue();
                }else{
                    nonAddressValue = config.getBlankValue();
                }
            }else {
                nonAddressValue = cellAddress.getPlaceholder();
            }
            unusedMap.put(cellAddress.getName(),nonAddressValue);
        }
//        for (String singleKey : onlyOnLeft.keySet()) {
//            if(referenceMapping.containsKey(singleKey)){
//                CellAddress cellAddress = referenceMapping.get(singleKey);
//                unusedMap.put(cellAddress.getName(),cellAddress.getDefaultValue());
//            }else{
//                unusedMap.put(singleKey,null);
//            }
//        }
        return unusedMap;
    }



    @SuppressWarnings("unchecked")
    private DesignConditions calculateConditions(List<?> circleDataList){
        DesignConditions designConditions = new DesignConditions();
        // 设置表索引
        int sheetIndex = context.getSwitchSheetIndex();
        designConditions.setSheetIndex(sheetIndex);
        // 判断是否是Map还是实体类并采集字段名
        Map<String, CellAddress> circleReferenceData = context.getCircleReferenceData().row(sheetIndex);
        Map<String, DesignConditions.FieldInfo> writeFieldNames = new HashMap<>();
        Object rowObjInstance = circleDataList.get(0);
        boolean ignoreException = config.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_EXCEPTION_RETURN_RESULT);
        if (rowObjInstance instanceof Map){
            designConditions.setSimplePOJO(false);
            Map<String, Object> rowObjInstanceMap = (Map<String, Object>) rowObjInstance;
            if (!rowObjInstanceMap.isEmpty()){
                writeFieldNames = rowObjInstanceMap.keySet()
                        .stream()
                        .collect(Collectors.toMap(key -> key, DesignConditions.FieldInfo::new));
            }
        }else {
            designConditions.setSimplePOJO(true);
            Class<?> instanceClass = rowObjInstance.getClass();
            boolean useGetter = config.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_USE_GETTER_METHOD);
            for (String key : circleReferenceData.keySet()) {
                AxolotlWriteIgnore ignore;
                Method getterMethod = null;
                Field field = ReflectToolkit.recursionGetField(instanceClass, key);
                DesignConditions.FieldInfo fieldInfo;
                if (useGetter){
                    String getterMethodName = ReflectToolkit.getFieldGetterMethodName(key);
                    fieldInfo = new DesignConditions.FieldInfo(getterMethodName);
                    fieldInfo.setGetter(true);
                    boolean isDirect = false;
                    try {
                        getterMethod = instanceClass.getMethod(getterMethodName);
                        AxolotlWriterGetter writerGetter = getterMethod.getAnnotation(AxolotlWriterGetter.class);
                        if (writerGetter == null){
                            if (field != null){
                                writerGetter = field.getAnnotation(AxolotlWriterGetter.class);
                            }
                        }
                        if (writerGetter != null){
                            LoggerHelper.debug(LOGGER, "[%s]模板重定向自定义Getter方法[%s]",key , writerGetter.value());
                            getterMethod = instanceClass.getMethod(writerGetter.value());
                            isDirect = true;
                        }
                    } catch (NoSuchMethodException noSuchMethodException) {
                        String[] methodNameSplit = noSuchMethodException.getMessage().split("\\.");
                        String methodName = methodNameSplit[methodNameSplit.length - 1];
                        methodName = methodName.substring(0, methodName.length() - 2);
                        LoggerHelper.error(LOGGER, "[%s]模板获取Getter方法[%s]失败",key , methodName);
                        getterMethod = null;
                        fieldInfo.setExist(false);
                        fieldInfo.setName(methodName);
                    }
                    if (getterMethod != null){
                        if (Modifier.isPublic(getterMethod.getModifiers())){
                            fieldInfo.setName(getterMethod.getName());
                            ignore = getterMethod.getAnnotation(AxolotlWriteIgnore.class);
                            if (ignore != null){
                                fieldInfo.setIgnore(true);
                            }else{
                                // 同名Getter可以忽略，非同名无法查找
                                if (!isDirect){
                                    if (field != null && field.getAnnotation(AxolotlWriteIgnore.class) != null){
                                        fieldInfo.setIgnore(true);
                                    }
                                }
                            }
                        }
                    }else{
                        if (ignoreException){
                            LoggerHelper.error(LOGGER, "未找到字段[%s]的Getter方法,写入将跳过该字段", key);
                            fieldInfo.setIgnore(true);
                            fieldInfo.setExist(false);
                        }else{
                            throw new AxolotlWriteException(LoggerHelper.format("未找到字段[%s]的Getter方法", key));
                        }
                    }
                    writeFieldNames.put(key,fieldInfo);
                }else{
                    fieldInfo = new DesignConditions.FieldInfo(key);
                    // 如果模板名称为Getter名称起始，优先寻找实体的Getter方法
                    if (key.startsWith(ReflectToolkit.GET_FIELD_LAMBDA) || key.startsWith(ReflectToolkit.IS_FIELD_LAMBDA)){
                        String tempName = key;
                        // 不为空说明有同名的字段,转为Getter方法名
                        if (field != null){
                            tempName = ReflectToolkit.getFieldGetterMethodName(tempName);
                        }
                        try {
                            getterMethod = instanceClass.getMethod(tempName);
                            if (Modifier.isPublic(getterMethod.getModifiers())){
                                fieldInfo.setName(tempName);
                                ignore = getterMethod.getAnnotation(AxolotlWriteIgnore.class);
                                if (ignore != null){
                                    LoggerHelper.debug(LOGGER, "Getter方法[%s]被忽略", tempName);
                                    fieldInfo.setIgnore(true);
                                }else{
                                    if (field != null){
                                        if (field.getAnnotation(AxolotlWriteIgnore.class) != null){
                                            LoggerHelper.debug(LOGGER, "字段[%s]被忽略", tempName);
                                            fieldInfo.setIgnore(true);
                                        }
                                    }else{
                                        fieldInfo.setGetter(true);
                                    }
                                }
                            }else{
                                LoggerHelper.debug(LOGGER, "Getter方法[%s]为私有,跳过使用", tempName);
                            }
                        } catch (NoSuchMethodException ignored) {
                            fieldInfo.setExist(false);
                        }
                        if (getterMethod == null && field != null){
                            if (field.getAnnotation(AxolotlWriteIgnore.class) != null){
                                LoggerHelper.debug(LOGGER, "字段[%s]被忽略,跳过使用", tempName);
                                fieldInfo.setIgnore(true);
                            }
                        }
                        writeFieldNames.put(key,fieldInfo);
                    }else{
                        if (field != null){
                            ignore = field.getAnnotation(AxolotlWriteIgnore.class);
                            if (ignore != null){
                                LoggerHelper.debug(LOGGER, "字段[%s]被忽略,跳过使用", key);
                                continue;
                            }
                            writeFieldNames.put(key, fieldInfo);
                        }
                    }
                }
            }
        }
        designConditions.setWriteFieldNames(writeFieldNames);
        ArrayList<String> writeFieldNamesList = new ArrayList<>(writeFieldNames.keySet());
        designConditions.setWriteFieldNamesList(writeFieldNamesList);
        // 判断字段第一次写入
        boolean initialWriting = context.fieldsIsInitialWriting(sheetIndex,writeFieldNamesList);
        // 添加写入字段记录
        context.addFieldRecords(sheetIndex,writeFieldNamesList, context.getCurrentWrittenBatch().get(sheetIndex));
        designConditions.setFieldsInitialWriting(initialWriting);
        // 漂移写入特性
        int startShiftRow = calculateStartShiftRow(circleReferenceData, designConditions, initialWriting);
        designConditions.setStartShiftRow(startShiftRow);
        // 获取非模板字段
        Map<String, CellAddress> nonWrittenAddress = findTemplateCell(initialWriting,
                startShiftRow, writeFieldNamesList,sheetIndex, circleReferenceData
        );
        designConditions.setNonWrittenAddress(nonWrittenAddress);
        designConditions.setNotTemplateCells(context.getSheetNonTemplateCells().get(sheetIndex, writeFieldNamesList));
        designConditions.setTemplateLineHeight(context.getLineHeightRecords().get(sheetIndex, writeFieldNamesList));
        return designConditions;
    }

    /**
     * 寻找非模板值
     *
     * @param initialWriting      是否初始写入
     * @param startShiftRow       起始行
     * @param writeFieldNamesList 写入字段名
     * @param sheetIndex          工作表索引
     * @param circleReferenceData 循环引用数据
     */
    private Map<String, CellAddress> findTemplateCell(boolean initialWriting,
                                                      int startShiftRow, List<String> writeFieldNamesList,
                                                      int sheetIndex, Map<String, CellAddress> circleReferenceData
    ){
        int templateLineIdx = initialWriting ? startShiftRow - 1 : startShiftRow;
        // 本行未在模板中的模板列（用于填充默认值或赋空值）
        Map<String,CellAddress> nonWrittenAddress = new HashMap<>();
        if(templateLineIdx < 0){
            return nonWrittenAddress;
        }
        Sheet sheet = getWorkbookSheet(sheetIndex);
        // 获取到写入字段的行次，并转为列和地址的映射
        Row templateRow = sheet.getRow(templateLineIdx);
        Map<Integer,CellAddress > templateColumnMap = circleReferenceData.values()
                .stream().filter(cellAddress -> cellAddress.getRowPosition() == templateLineIdx)
                .collect(Collectors.toMap(CellAddress::getColumnPosition, cellAddress -> cellAddress));

        boolean alreadyFill = context.getSheetNonTemplateCells().contains(sheetIndex, writeFieldNamesList);
        boolean nonTemplateCellFill = config.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL);

        // 模板行种非模板列
        List<CellAddress> nonTemplateCellAddressList = new ArrayList<>();
        for (int i = 0; i < templateRow.getLastCellNum(); i++) {
            if(!templateColumnMap.containsKey(i)){
                if (alreadyFill){continue;}
                // 将非模板列存储
                if(initialWriting && nonTemplateCellFill){
                    Cell cell = templateRow.getCell(i);
                    if(cell == null){continue;}
                    CellAddress nonTempalteCellAddress = new CellAddress(null, templateLineIdx, i, cell.getCellStyle());
                    nonTempalteCellAddress.set_nonTemplateCell(cell);
                    nonTempalteCellAddress.setMergeRegion(ExcelToolkit.isCellMerged(sheet, templateLineIdx, i));
                    nonTemplateCellAddressList.add(nonTempalteCellAddress);
                }
            }else{
                // 本次未写入的地址
                CellAddress cellAddress = templateColumnMap.get(i);
                String name = cellAddress.getName();
                if (!writeFieldNamesList.contains(name)) {nonWrittenAddress.put(name,cellAddress);}
            }
        }
        // 存储非模板列
        if(!alreadyFill){
            LoggerHelper.debug(LOGGER,"获取模板行[%s]个非模板列",nonTemplateCellAddressList.size());
            context.getSheetNonTemplateCells().put(sheetIndex,writeFieldNamesList,nonTemplateCellAddressList);
        }
        return nonWrittenAddress;
    }

    /**
     * 写入占位符列表数据到工作表
     *
     * @param sheet 工作表
     * @param circleDataList 循环列表数据
     */
    @SneakyThrows
    @SuppressWarnings("unchecked")
    private void writeCircleData(Sheet sheet, List<?> circleDataList){
        // 数据不为空则写入数据
        if (Validator.objNotNull(circleDataList)){
            DesignConditions designConditions = this.calculateConditions(circleDataList);
            LoggerHelper.debug(LOGGER,"本次写入字段为:%s",designConditions.getWriteFieldNamesList());
            boolean initialWriting = designConditions.isFieldsInitialWriting();
            int startShiftRow = designConditions.getStartShiftRow();
            if ((circleDataList.size() > 1 || (circleDataList.size() == 1 && initialWriting)) &&
                    config.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW)){
                // 最后一行大于起始行，则下移，否则为表底不下移
                int lastRowNum = sheet.getLastRowNum();
                if(startShiftRow >= 0 && lastRowNum >= startShiftRow){
                    int shiftRowNumber = initialWriting ? circleDataList.size() - 1 : circleDataList.size();
                    LoggerHelper.debug(LOGGER,"当前写入起始行次[%s],下移行次:[%s],",startShiftRow,shiftRowNumber);
                    if (shiftRowNumber > 0){
                        sheet.shiftRows(startShiftRow, sheet.getLastRowNum(), shiftRowNumber, true,true);
                    }
                }
            }
            int sheetIndex = designConditions.getSheetIndex();
            Map<String, CellAddress> circleReferenceData = context.getCircleReferenceData().row(sheetIndex);
            // 写入列表数据
            HashBasedTable<Integer, String, Boolean> alreadyUsedReferenceData = context.getAlreadyUsedReferenceData();
            Map<String, Boolean> alreadyUsedDataMapping = alreadyUsedReferenceData.row(context.getSwitchSheetIndex());
            Map<String, CellAddress> calculateReferenceData = this.context.getCalculateReferenceData().row(context.getSwitchSheetIndex());
            Map<String, DesignConditions.FieldInfo> writeFieldNames = designConditions.getWriteFieldNames();
            boolean ignoreException = config.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_EXCEPTION_RETURN_RESULT);
            boolean useDictCode = config.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_USE_DICT_CODE_TRANSFER);
            for (Object data : circleDataList) {
                HashSet<Integer> alreadyWrittenMergeRegionColumns = new HashSet<>();
                boolean isCurrentBatchData = false;
                boolean alreadySetLineHeight = false;
                for (Map.Entry<String, CellAddress> fieldMapping : circleReferenceData.entrySet()) {
                    String fieldMappingKey = fieldMapping.getKey();
                    CellAddress cellAddress = circleReferenceData.get(fieldMappingKey);
                    int rowPosition = cellAddress.getRowPosition();
                    if(!alreadySetLineHeight){
                        Row templateRow = ExcelToolkit.createOrCatchRow(sheet, rowPosition);
                        Short templateLineHeight = designConditions.getTemplateLineHeight();
                        if (templateLineHeight != null && templateLineHeight != -1) {
                            templateRow.setHeight(templateLineHeight);
                        }
                        alreadySetLineHeight = true;
                    }
                    // 判断是否已经写入
                    boolean isWritten = false;
                    if(writeFieldNames.containsKey(fieldMappingKey)){
                        Object value;
                        DesignConditions.FieldInfo fieldInfo = writeFieldNames.get(fieldMappingKey);
                        if (fieldInfo.isIgnore() || !fieldInfo.isExist()){
                            value = null;
                        }else{
                            if (designConditions.isSimplePOJO()){
                                Class<?> dataClass = data.getClass();
                                if (fieldInfo.isGetter()){
                                    Method method = dataClass.getMethod(fieldInfo.getName());
                                    try {
                                        value = method.invoke(data);
                                    }catch (Exception exception){
                                        if (ignoreException){
                                            LoggerHelper.error(LOGGER,"获取字段[%s]值失败,将赋予空值",fieldMappingKey);
                                            value = null;
                                        }else{throw exception;}
                                    }
                                }else{
                                    Field field = dataClass.getDeclaredField(fieldMappingKey);
                                    field.setAccessible(true);
                                    value = field.get(data);
                                }
                            }else{
                                Map<String, Object> map = (Map<String, Object>) data;
                                value = map.get(fieldMappingKey);
                            }
                        }
                        //设置单元格值
                        Cell writableCell = ExcelToolkit.createOrCatchCell(sheet, rowPosition, cellAddress.getColumnPosition(), cellAddress.getCellStyle());
                        // 空值时使用默认值填充
                        if (Validator.strIsBlank(value)){
                            String defaultValue = cellAddress.getDefaultValue();
                            if (defaultValue != null){
                                writableCell.setCellValue(cellAddress.replacePlaceholder(defaultValue));
                            }else{
                                if (config.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL)){
                                    writableCell.setCellValue(cellAddress.replacePlaceholder(config.getBlankValue()));
                                    isWritten = true;
                                }else{
                                    writableCell.setBlank();
                                }
                            }
                        }else {
                            // 计算写入列的值
                            if (calculateReferenceData.containsKey(fieldMappingKey) && Validator.strIsNumber(value.toString())){
                                CellAddress calculateAddress = calculateReferenceData.get(fieldMappingKey);
                                BigDecimal calculatedValue = calculateAddress.getCalculatedValue();
                                calculatedValue = calculatedValue.add(new BigDecimal(value.toString()));
                                calculateAddress.setCalculatedValue(calculatedValue);
                            }
                            // 暂时只适配String类型
                            String valueString = config.getDataInverter().convert(value).toString();
                            if (useDictCode){
                                valueString = componentRender.convertDictCodeToName(sheetIndex,String.class,fieldMappingKey,data,valueString);
                            }
                            writableCell.setCellValue(cellAddress.replacePlaceholder(valueString));
                            isWritten = true;
                        }
                    }else if (designConditions.getNonWrittenAddress().containsKey(fieldMappingKey)){
                        cellAddress = designConditions.getNonWrittenAddress().get(fieldMappingKey);
                        Object nonAddressValue;
                        if (config.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_PLACEHOLDER_FILL_DEFAULT)){
                            if(cellAddress.getDefaultValue() != null){
                                nonAddressValue = cellAddress.getDefaultValue();
                            }else{
                                nonAddressValue = config.getBlankValue();
                            }
                        }else {
                            nonAddressValue = cellAddress.getCellValue();
                        }
                        ExcelToolkit.cellAssignment(
                                sheet, rowPosition, cellAddress.getColumnPosition(),
                                cellAddress.getCellStyle(), nonAddressValue
                        );
                        isWritten = true;
                    }
                    if (isWritten){
                        this.setMergeRegion(sheet,cellAddress,rowPosition,alreadyWrittenMergeRegionColumns);
                        cellAddress.setRowPosition(++rowPosition);
                        if (!alreadyUsedDataMapping.containsKey(cellAddress.getPlaceholder())){
                            alreadyUsedDataMapping.put(cellAddress.getPlaceholder(),true);
                        }
                        isCurrentBatchData = true;
                    }
                }
                // 填充非模板单元格
                if (isCurrentBatchData && config.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL)){
                    List<CellAddress> nonTemplateCells = context.getSheetNonTemplateCells().get(sheetIndex, designConditions.getWriteFieldNamesList());
                    if (nonTemplateCells != null && !nonTemplateCells.isEmpty()) {
                        for (CellAddress nonTemplateCellAddress : nonTemplateCells){
                            Cell nonTemplateCell = nonTemplateCellAddress.get_nonTemplateCell();
                            int rowPosition = nonTemplateCellAddress.getRowPosition();
                            if(!nonTemplateCellAddress.isInitializedWrite()){
                                if(nonTemplateCellAddress.isMergeCell()){
                                    int columnIndex = nonTemplateCell.getColumnIndex();
                                    if(alreadyWrittenMergeRegionColumns.contains(columnIndex)){
                                        continue;
                                    }else{
                                        this.setMergeRegion(sheet,nonTemplateCellAddress,rowPosition,alreadyWrittenMergeRegionColumns);
                                    }
                                }
                                Cell writableCell = ExcelToolkit.createOrCatchCell(sheet, rowPosition, nonTemplateCellAddress.getColumnPosition(), null);
                                ExcelToolkit.cloneOldCell2NewCell(writableCell,nonTemplateCell);
                            }
                            nonTemplateCellAddress.setRowPosition(++rowPosition);
                        }
                    }
                }
            }
        }
    }



    /**
     * 设置合并单元格区域
     * @param sheet 工作表
     * @param cellAddress 单元格地址信息
     * @param rowPosition 行次
     */
    private void setMergeRegion(Sheet sheet,CellAddress cellAddress,int rowPosition,HashSet<Integer> alreadyWrittenMergeRegionColumns){
        if (cellAddress.isMergeCell()){
            CellRangeAddress mergeRegion = cellAddress.getMergeRegion();
            for (int i = mergeRegion.getFirstColumn(); i <= mergeRegion.getLastColumn(); i++) {
                alreadyWrittenMergeRegionColumns.add(i);
            }
            if(!cellAddress.isInitializedWrite()){
                mergeRegion.setFirstRow(rowPosition);
                mergeRegion.setLastRow(rowPosition);
                StyleHelper.renderMergeRegionStyle(sheet,mergeRegion,cellAddress.getCellStyle());
                sheet.addMergedRegion(mergeRegion);
            }
        }
    }

    /**
     * 计算起始行
     * @param circleReferenceData 引用数据
     * @param designConditions 参数条件
     * @param initialWriting 是否是第一次写入
     * @return 起始行
     */
    private int calculateStartShiftRow(Map<String, CellAddress> circleReferenceData, DesignConditions designConditions, boolean initialWriting) {
        int maxRowPosition = Integer.MIN_VALUE;
        Map<String, DesignConditions.FieldInfo> writeFieldNames = designConditions.getWriteFieldNames();
        for (Map.Entry<String, CellAddress> addressEntry : circleReferenceData.entrySet()) {
            if (writeFieldNames.containsKey(addressEntry.getKey())){
                maxRowPosition = Math.max(maxRowPosition, addressEntry.getValue().getRowPosition());
            }
        }
        if (initialWriting){
            int sheetIndex = designConditions.getSheetIndex();
            if(maxRowPosition >= 0){
                // 设置模板行行高
                Sheet sheet = this.getWorkbookSheet(sheetIndex);
                short templateRowHeight = sheet.getRow(maxRowPosition).getHeight();
                LoggerHelper.debug(LOGGER,"设置模板行[%s]行高为[%s]",maxRowPosition,templateRowHeight);
                context.getLineHeightRecords().put(sheetIndex, designConditions.getWriteFieldNamesList(),templateRowHeight);
            }else{
                context.getLineHeightRecords().put(sheetIndex, designConditions.getWriteFieldNamesList(), (short) -1);
                LoggerHelper.debug(LOGGER,"未找到任意占位符,取消设置行高.");
            }
        }
        // 第一次写入需要跳过占位符那一行，所以移动需要少一行
        return initialWriting ? maxRowPosition + 1 : maxRowPosition;
    }

    /**
     * 注入内置变量
     *
     * @param singleMap 单条数据
     * @param gatherUnusedStage 是否收集未使用的占位符
     */
    @SuppressWarnings({"rawtypes","unchecked"})
    private void injectCommonConstInfo(Map singleMap, boolean gatherUnusedStage){
        if(!gatherUnusedStage && context.isFirstBatch(context.getSwitchSheetIndex())){
            if (singleMap == null){
                singleMap = new HashMap<>();
            }
            singleMap.put(AxolotlConstant.CREATE_TIME, Time.getCurrentTime());
            singleMap.put(AxolotlConstant.CREATE_DATE, Time.regexTime(Time.YMD_HORIZONTAL_FORMAT_REGEX,new Date()));
            LoggerHelper.debug(LOGGER, "注入内置常量");
        }
    }

    /**
     * 解析模板占位符
     *
     * @param sheet 工作表
     */
    private void resolveTemplate(Sheet sheet,boolean isFinal){
        int sheetIndex = getSheetIndex(sheet);;
        if (!context.getResolvedSheetRecord().containsKey(sheetIndex) || isFinal){
            HashBasedTable<Integer, String, CellAddress> singleReferenceData = context.getSingleReferenceData();
            HashBasedTable<Integer, String, CellAddress> circleReferenceData = context.getCircleReferenceData();
            HashBasedTable<Integer, String, CellAddress> calculateReferenceData = context.getCalculateReferenceData();
            for (int rowIdx = 0; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
                Row row = sheet.getRow(rowIdx);
                if (row != null){
                    short lastCellNum = row.getLastCellNum();
                    for (int colIdx = 0; colIdx < lastCellNum; colIdx++) {
                        Cell cell = row.getCell(colIdx);
                        // 占位符必然是字符串类型
                        if (cell != null && CellType.STRING.equals(cell.getCellType())){
                            Boolean foundPlaceholder = findPlaceholderData(isFinal,singleReferenceData,
                                    TemplatePlaceholderPattern.SINGLE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cell);
                            if (!foundPlaceholder){
                                foundPlaceholder = findPlaceholderData(isFinal,circleReferenceData,
                                        TemplatePlaceholderPattern.CIRCLE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cell);
                            }
                            if (!foundPlaceholder) {
                                findPlaceholderData(isFinal,calculateReferenceData, TemplatePlaceholderPattern.AGGREGATE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cell);
                            }
                        }
                    }
                }
            }
            int singleReferenceDataSize = context.getSingleReferenceData().size();
            int circleReferenceDataSize = context.getCircleReferenceData().size();
            int calculateReferenceDataSize = context.getCalculateReferenceData().size();
            context.getResolvedSheetRecord().put(sheetIndex,true);
            LoggerHelper.debug(LOGGER, LoggerHelper.format("%s工作表索引[%s]解析模板完成，共解析到[%s]个占位符,引用占位符[%s]个,列表占位符[%s]个,计算占位符[%s]个",
                    isFinal? "[收尾阶段]":"",
                    sheetIndex,
                    singleReferenceDataSize + circleReferenceDataSize + calculateReferenceDataSize,
                    singleReferenceDataSize,
                    circleReferenceDataSize,
                    calculateReferenceDataSize
            ));
        }else{
            LoggerHelper.debug(LOGGER, LoggerHelper.format("工作表[%s]已被解析过，跳过本次解析",sheetIndex));
        }
    }

    /**
     * 解析模板值到变量
     * 模板值为两部分组成：
     * ${name:111}
     * 一个是占位符本身name，另一个是默认值111，由:分隔
     * @param isFinal 是否是收尾阶段
     * @param referenceData 引用数据
     * @param pattern       模板匹配正则
     * @param sheetIndex    工作簿索引
     * @param cell   当前单元格
     */
    private Boolean findPlaceholderData(boolean isFinal, HashBasedTable<Integer, String, CellAddress> referenceData,
                                        Pattern pattern, int sheetIndex,Cell cell) {
        List<CellAddress> cellAddressList = new ArrayList<>();
        Map<String, CellAddress> cellAddressMap = referenceData.row(sheetIndex);
        int cellMultipleMatchTemplate = -1;
        String stringCellValue = cell.getStringCellValue();
        Matcher matcher = pattern.matcher(stringCellValue);
        DataFormat dataFormat = workbook.createDataFormat();
        short textFormatIndex = dataFormat.getFormat("@");
        while (matcher.find()){
            cellMultipleMatchTemplate++;
            int rowIndex = cell.getRowIndex();
            int columnIndex = cell.getColumnIndex();
            CellStyle cellStyle = cell.getCellStyle();
            cellStyle.setDataFormat(textFormatIndex);
            CellAddress cellAddress = new CellAddress(stringCellValue,rowIndex ,columnIndex , cellStyle);
            cellAddress.setPlaceholder(matcher.group());
            String matchTemplate = matcher.group(1);
            String[] defaultSplitContent = matchTemplate.split(StringPool.COLON);
            String name = defaultSplitContent[0];
            cellAddress.setName(name);
            if (defaultSplitContent.length > 1){
                cellAddress.setDefaultValue(defaultSplitContent[1]);
            }
            CellRangeAddress cellMerged = ExcelToolkit.isCellMerged(getWorkbookSheet(sheetIndex), rowIndex, columnIndex);
            if (cellMerged != null){
                LoggerHelper.debug(LOGGER, LoggerHelper.format("解析到占位符[%s]为合并单元格[%s]",cellAddress.getPlaceholder(),cellMerged.formatAsString()));
                cellAddress.setMergeRegion(cellMerged);
            }
            boolean isCirclePattern = pattern.equals(TemplatePlaceholderPattern.CIRCLE_REFERENCE_TEMPLATE_PATTERN);
            if (isCirclePattern || pattern.equals(TemplatePlaceholderPattern.SINGLE_REFERENCE_TEMPLATE_PATTERN)){
                cellAddress.setPlaceholderType(isCirclePattern ? PlaceholderType.CIRCLE : PlaceholderType.MAPPING);
                cellAddressMap.put(name, cellAddress);
            }else if (pattern.equals(TemplatePlaceholderPattern.AGGREGATE_REFERENCE_TEMPLATE_PATTERN)){
                if (!isFinal){
                    cellAddress.setPlaceholderType(PlaceholderType.CALCULATE);
                    cellAddress.setCalculatedValue(BigDecimal.ZERO);
                    cellAddressMap.put(name, cellAddress);
                }else {
                    if (cellAddressMap.containsKey(name)){
                        CellAddress originalAddress = cellAddressMap.get(name);
                        originalAddress.setRowPosition(cellAddress.getRowPosition());
                        cellAddressMap.put(name, originalAddress);
                    }else{
                        cellAddressMap.put(name, cellAddress);
                    }
                }
            }
            cellAddressList.add(cellAddress);
        }
        if (cellMultipleMatchTemplate > 0){
            for (CellAddress cellAddress : cellAddressList) {
                cellAddress.setCellMultipleMatchTemplate(true);
            }
        }
        return !cellAddressList.isEmpty();
    }

    /**
     * 写入器刷新内容
     * 进入写入剩余内容进入关闭流前的收尾工作
     *
     * @param isFinal 是否是最终刷新，关闭写入前的最后一次刷新
     */
    public void flush(boolean isFinal) {
        if (isFinal){
            for (Integer i : context.getResolvedSheetRecord().keySet()) {
                this.context.setSwitchSheetIndex(i);
                this.resolveTemplate(getWorkbookSheet(i), true);
                this.gatherUnusedSingleReferenceDataAndFillDefault();
                this.gatherUnusedCircleReferenceDataAndFillDefault();
                this.setCalculateData(i);
            }
        }else{
            this.resolveTemplate(getWorkbookSheet(context.getSwitchSheetIndex()),false);
        }
    }

    @Override
    public void flush() {
        // 采集未映射数据
        this.flush(false);
    }

    /**
     * 关闭工作簿所对应输出流
     */
    @SneakyThrows
    @Override
    public void close() {
        LoggerHelper.debug(LOGGER, "工作薄写入进入关闭阶段");
        OutputStream outputStream = config.getOutputStream();
        if(outputStream != null){
            this.flush(true);
            workbook.write(config.getOutputStream());
            workbook.close();
            config.close();
        }else{
            String message = "输出流为空,请指定输出流";
            LoggerHelper.debug(LOGGER,message);
            throw new AxolotlWriteException(message);
        }
    }

    @Override
    public void switchSheet(int sheetIndex) {
        super.switchSheet(sheetIndex);
        this.resolveTemplate(this.getWorkbookSheet(sheetIndex),false);
    }
}

package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.annotations.AxolotlWriteIgnore;
import cn.toutatis.xvoid.axolotl.excel.writer.components.annotations.AxolotlWriterGetter;
import cn.toutatis.xvoid.axolotl.excel.writer.constant.TemplatePlaceholderPattern;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.*;
import cn.toutatis.xvoid.axolotl.exceptions.AxolotlException;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.axolotl.toolkit.tika.TikaShell;
import cn.toutatis.xvoid.common.standard.StringPool;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import com.alibaba.fastjson.JSON;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.MapDifference;
import com.google.common.collect.Maps;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

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
    private final TemplateWriteConfig writeConfig;

    private final TemplateWriteContext writeContext;

    /**
     * 主构造函数
     *
     * @param templateWriteConfig 写入配置
     */
    public AxolotlTemplateExcelWriter(TemplateWriteConfig templateWriteConfig) {
        super.LOGGER = LOGGER;
        this.writeConfig = templateWriteConfig;
        this.checkWriteConfig(this.writeConfig);
        TemplateWriteContext templateWriteContext = new TemplateWriteContext();
        super.writeContext = templateWriteContext;
        this.writeContext = templateWriteContext;
        this.writeContext.setSwitchSheetIndex(writeConfig.getSheetIndex());
    }

    /**
     * 构造函数
     * 可以写入一个模板文件
     *
     * @param templateFile 模板文件
     * @param writeConfig 写入配置
     */
    public AxolotlTemplateExcelWriter(File templateFile, TemplateWriteConfig writeConfig) {
        this(writeConfig);
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
        LoggerHelper.info(LOGGER, writeContext.getCurrentWrittenBatchAndIncrement(writeContext.getSwitchSheetIndex()));
        XSSFSheet sheet;
        // 判断是否是模板写入
        AxolotlWriteResult axolotlWriteResult = new AxolotlWriteResult();
        if (writeContext.isTemplateWrite()){
            int switchSheetIndex = writeContext.getSwitchSheetIndex();
            sheet = getWorkbookSheet(switchSheetIndex);
            // 只有第一次写入时解析模板占位符
            if (writeContext.isFirstBatch(switchSheetIndex)){
                // 解析模板占位符到上下文
                this.resolveTemplate(sheet,false);
            }
//            // 写入Map映射
            this.writeSingleData(sheet, fixMapping,writeContext.getSingleReferenceData(),false);
//            // 写入循环数据
            this.writeCircleData(sheet, dataList);
            axolotlWriteResult.setWrite(true);
            axolotlWriteResult.setMessage("写入完成");
        }else{
            String message = "非模板写入请使用AxolotlAutoExcelWriter.write()方法";
            if(writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_EXCEPTION_RETURN_RESULT)){
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
        int sheetIndex = workbook.getXSSFWorkbook().getSheetIndex(sheet);
        Map<String, CellAddress> addressMapping = referenceData.row(sheetIndex);
        // 记录已使用的引用数据
        Map<String, Boolean> alreadyUsedReferenceData = writeContext.getAlreadyUsedReferenceData().row(sheetIndex);
        for (String singleKey : addressMapping.keySet()) {
            // 如果地址引用包含该关键字，则写入数据
            CellAddress cellAddress = addressMapping.get(singleKey);
            String placeholder = cellAddress.getPlaceholder();
            if(dataMapping.containsKey(singleKey)){
                // 已经写入过则跳过写入
                if(alreadyUsedReferenceData.containsKey(placeholder)){
                    debug(LOGGER, format("已跳过使用的占位符[%s]",placeholder));
                    continue;
                }
                // 写入单元格值
                Cell cell = sheet.getRow(cellAddress.getRowPosition()).getCell(cellAddress.getColumnPosition());
                String stringCellValue = cell.getStringCellValue();
                Object info = dataMapping.get(singleKey);
                if(info != null){
                    debug(LOGGER, format("设置模板占位符[%s]值[%s]",placeholder,info));
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
                        debug(LOGGER, format("%s设置模板占位符[%s]为空值",gatherUnusedStage ? "[收尾阶段]":"",placeholder));
                        cell.setBlank();
                    }else{
                        debug(LOGGER, format("%s设置模板占位符[%s]为%s值",gatherUnusedStage ? "[收尾阶段]":"", placeholder,isDefaultValue ? "默认":"空"));
                        cell.setCellValue(newCellValue);
                    }
                }
                cellAddress.setWrittenRow(cell.getRowIndex());
                alreadyUsedReferenceData.put(placeholder,true);
            }else{
                debug(LOGGER, format("未使用模板占位符[%s]",placeholder));
            }
        }
    }

    /**
     * 未使用的单次占位符填充默认值
     */
    private void gatherUnusedSingleReferenceDataAndFillDefault() {
        if(writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_PLACEHOLDER_FILL_DEFAULT)){
            int sheetIndex = writeContext.getSwitchSheetIndex();
            Sheet sheet = this.getWorkbookSheet(writeContext.getSwitchSheetIndex());
            Map<String, CellAddress> singleReferenceMapping =  writeContext.getSingleReferenceData().row(sheetIndex);
            HashMap<String, Object> unusedMap = gatherUnusedField(sheetIndex, singleReferenceMapping);
            this.writeSingleData(sheet,unusedMap,writeContext.getSingleReferenceData(),true);
        }
    }

    /**
     * 未使用的列表占位符填充默认值
     */
    private void gatherUnusedCircleReferenceDataAndFillDefault() {
        if(writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_PLACEHOLDER_FILL_DEFAULT)){
            int sheetIndex = writeContext.getSwitchSheetIndex();
            Sheet sheet = this.getWorkbookSheet(sheetIndex);
            Map<String, CellAddress> circleReferenceData =  writeContext.getCircleReferenceData().row(sheetIndex);
            HashMap<String, Object> map = gatherUnusedField(sheetIndex, circleReferenceData);
            this.writeSingleData(sheet,map,writeContext.getCircleReferenceData(),true);
        }
    }

    /**
     * 设置计算数据
     */
    private void setCalculateData(int sheetIndex) {
        HashBasedTable<Integer, String, CellAddress> calculateReferenceData = writeContext.getCalculateReferenceData();
        Map<String, CellAddress> calculateData = calculateReferenceData.row(sheetIndex);
        Map<String, BigDecimal> extract = new HashMap<>();
        for (String s : calculateData.keySet()) {
            BigDecimal calculatedValue = calculateData.get(s).getCalculatedValue();
            extract.put(s,calculatedValue.setScale(AxolotlDefaultReaderConfig.XVOID_DEFAULT_DECIMAL_SCALE, RoundingMode.HALF_UP));
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
        Map<String, Boolean> alreadyUsedDataMapping =  writeContext.getAlreadyUsedReferenceData().row(sheetIndex);
        MapDifference<String, Object> difference = Maps.difference(referenceMapping, alreadyUsedDataMapping);
        Map<String, Object> onlyOnLeft = difference.entriesOnlyOnLeft();
        HashMap<String, Object> unusedMap = new HashMap<>();
        for (String singleKey : onlyOnLeft.keySet()) {
            if(referenceMapping.containsKey(singleKey)){
                unusedMap.put(singleKey,referenceMapping.get(singleKey).getDefaultValue());
            }else{
                unusedMap.put(singleKey,null);
            }
        }
        return unusedMap;
    }



    @SuppressWarnings("unchecked")
    private DesignConditions calculateConditions(List<?> circleDataList){
        DesignConditions designConditions = new DesignConditions();
        // 设置表索引
        int sheetIndex = writeContext.getSwitchSheetIndex();
        designConditions.setSheetIndex(sheetIndex);
        // 判断是否是Map还是实体类并采集字段名
        Map<String, CellAddress> circleReferenceData = writeContext.getCircleReferenceData().row(sheetIndex);
        Map<String, DesignConditions.FieldInfo> writeFieldNames = new HashMap<>();
        Object rowObjInstance = circleDataList.get(0);
        boolean ignoreException = writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_EXCEPTION_RETURN_RESULT);
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
            boolean useGetter = writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_USE_GETTER_METHOD);
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
                            debug(LOGGER, "[%s]模板重定向自定义Getter方法[%s]",key , writerGetter.value());
                            getterMethod = instanceClass.getMethod(writerGetter.value());
                            isDirect = true;
                        }
                    } catch (NoSuchMethodException noSuchMethodException) {
                        String[] methodNameSplit = noSuchMethodException.getMessage().split("\\.");
                        String methodName = methodNameSplit[methodNameSplit.length - 1];
                        methodName = methodName.substring(0, methodName.length() - 2);
                        error(LOGGER, "[%s]模板获取Getter方法[%s]失败",key , methodName);
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
                            error(LOGGER, "未找到字段[%s]的Getter方法,写入将跳过该字段", key);
                            fieldInfo.setIgnore(true);
                            fieldInfo.setExist(false);
                        }else{
                            throw new AxolotlWriteException(format("未找到字段[%s]的Getter方法", key));
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
                                    debug(LOGGER, "Getter方法[%s]被忽略", tempName);
                                    fieldInfo.setIgnore(true);
                                }else{
                                    if (field != null){
                                        if (field.getAnnotation(AxolotlWriteIgnore.class) != null){
                                            debug(LOGGER, "字段[%s]被忽略", tempName);
                                            fieldInfo.setIgnore(true);
                                        }
                                    }else{
                                        fieldInfo.setGetter(true);
                                    }
                                }
                            }else{
                                debug(LOGGER, "Getter方法[%s]为私有,跳过使用", tempName);
                            }
                        } catch (NoSuchMethodException ignored) {
                            fieldInfo.setExist(false);
                        }
                        if (getterMethod == null && field != null){
                            if (field.getAnnotation(AxolotlWriteIgnore.class) != null){
                                debug(LOGGER, "字段[%s]被忽略,跳过使用", tempName);
                                fieldInfo.setIgnore(true);
                            }
                        }
                        writeFieldNames.put(key,fieldInfo);
                    }else{
                        if (field != null){
                            ignore = field.getAnnotation(AxolotlWriteIgnore.class);
                            if (ignore != null){
                                debug(LOGGER, "字段[%s]被忽略,跳过使用", key);
                                continue;
                            }
                            writeFieldNames.put(key, fieldInfo);
                        }
                    }
                }
            }
        }
        System.err.println(JSON.toJSONString(writeFieldNames));
        designConditions.setWriteFieldNames(writeFieldNames);
        ArrayList<String> writeFieldNamesList = new ArrayList<>(writeFieldNames.keySet());
        designConditions.setWriteFieldNamesList(writeFieldNamesList);
        // 判断字段第一次写入
        boolean initialWriting = writeContext.fieldsIsInitialWriting(sheetIndex,writeFieldNamesList);
        // 添加写入字段记录
        writeContext.addFieldRecords(sheetIndex,writeFieldNamesList,writeContext.getCurrentWrittenBatch().get(sheetIndex));
        designConditions.setFieldsInitialWriting(initialWriting);
        // 漂移写入特性
        int startShiftRow = calculateStartShiftRow(circleReferenceData, designConditions, initialWriting);
        designConditions.setStartShiftRow(startShiftRow);
        // 获取非模板字段
        Map<String, CellAddress> nonWrittenAddress = findTemplateCell(initialWriting,
                startShiftRow, writeFieldNamesList,sheetIndex, circleReferenceData
        );
        designConditions.setNonWrittenAddress(nonWrittenAddress);
        designConditions.setNotTemplateCells(writeContext.getSheetNonTemplateCells().get(sheetIndex, writeFieldNamesList));
        designConditions.setTemplateLineHeight(writeContext.getLineHeightRecords().get(sheetIndex, writeFieldNamesList));
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
        XSSFSheet sheet = getWorkbookSheet(sheetIndex);
        // 获取到写入字段的行次，并转为列和地址的映射
        XSSFRow templateRow = sheet.getRow(templateLineIdx);
        Map<Integer,CellAddress > templateColumnMap = circleReferenceData.values()
                .stream().filter(cellAddress -> cellAddress.getRowPosition() == templateLineIdx)
                .collect(Collectors.toMap(CellAddress::getColumnPosition, cellAddress -> cellAddress));

        boolean alreadyFill = writeContext.getSheetNonTemplateCells().contains(sheetIndex, writeFieldNamesList);
        boolean nonTemplateCellFill = writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL);

        // 模板行种非模板列
        List<CellAddress> nonTemplateCellAddressList = new ArrayList<>();
        for (int i = 0; i < templateRow.getLastCellNum(); i++) {
            if(!templateColumnMap.containsKey(i)){
                if (alreadyFill){continue;}
                // 将非模板列存储
                if(initialWriting && nonTemplateCellFill){
                    XSSFCell cell = templateRow.getCell(i);
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
            debug(LOGGER,"获取模板行[%s]个非模板列",nonTemplateCellAddressList.size());
            writeContext.getSheetNonTemplateCells().put(sheetIndex,writeFieldNamesList,nonTemplateCellAddressList);
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
    private void writeCircleData(XSSFSheet sheet, List<?> circleDataList){
        // 数据不为空则写入数据
        if (Validator.objNotNull(circleDataList)){
            DesignConditions designConditions = this.calculateConditions(circleDataList);
            LoggerHelper.debug(LOGGER,"本次写入字段为:%s",designConditions.getWriteFieldNamesList());
            boolean initialWriting = designConditions.isFieldsInitialWriting();
            int startShiftRow = designConditions.getStartShiftRow();
            if ((circleDataList.size() > 1 || (circleDataList.size() == 1 && initialWriting)) &&
                    writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW)){
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
            Map<String, CellAddress> circleReferenceData = writeContext.getCircleReferenceData().row(sheetIndex);
            // 写入列表数据
            HashBasedTable<Integer, String, Boolean> alreadyUsedReferenceData = writeContext.getAlreadyUsedReferenceData();
            Map<String, Boolean> alreadyUsedDataMapping = alreadyUsedReferenceData.row(writeContext.getSwitchSheetIndex());
            Map<String, CellAddress> calculateReferenceData = this.writeContext.getCalculateReferenceData().row(writeContext.getSwitchSheetIndex());
            Map<String, DesignConditions.FieldInfo> writeFieldNames = designConditions.getWriteFieldNames();
            boolean ignoreException = writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_EXCEPTION_RETURN_RESULT);
            for (Object data : circleDataList) {
                HashSet<Integer> alreadyWrittenMergeRegionColumns = new HashSet<>();
//                debug(LOGGER,"[写入数据]"+data);
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
                                            value = null;
                                        }else{
                                            throw exception;
                                        }
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
                        // TODO 转换字典
                        Cell writableCell = ExcelToolkit.createOrCatchCell(sheet, rowPosition, cellAddress.getColumnPosition(), cellAddress.getCellStyle());
                        // 空值时使用默认值填充
                        if (Validator.strIsBlank(value)){
                            String defaultValue = cellAddress.getDefaultValue();
                            if (defaultValue != null){
                                writableCell.setCellValue(cellAddress.replacePlaceholder(defaultValue));
                            }else{
                                if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL)){
                                    writableCell.setCellValue(cellAddress.replacePlaceholder(""));
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
                            writableCell.setCellValue(cellAddress.replacePlaceholder(writeConfig.getDataInverter().convert(value).toString()));
                        }
                        isWritten = true;
                    }else if (designConditions.getNonWrittenAddress().containsKey(fieldMappingKey)){
                        ExcelToolkit.cellAssignment(
                                sheet, rowPosition, cellAddress.getColumnPosition(),
                                cellAddress.getCellStyle(), cellAddress.getDefaultValue()
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
                if (isCurrentBatchData && writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL)){
                    List<CellAddress> nonTemplateCells = writeContext.getSheetNonTemplateCells().get(sheetIndex, designConditions.getWriteFieldNamesList());
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
                XSSFSheet sheet = this.getWorkbookSheet(sheetIndex);
                short templateRowHeight = sheet.getRow(maxRowPosition).getHeight();
                debug(LOGGER,"设置模板行[%s]行高为[%s]",maxRowPosition,templateRowHeight);
                writeContext.getLineHeightRecords().put(sheetIndex, designConditions.getWriteFieldNamesList(),templateRowHeight);
            }else{
                writeContext.getLineHeightRecords().put(sheetIndex, designConditions.getWriteFieldNamesList(), (short) -1);
                debug(LOGGER,"未找到任意占位符,取消设置行高.");
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
        if(!gatherUnusedStage && writeContext.isFirstBatch(writeContext.getSwitchSheetIndex())){
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
        int sheetIndex = workbook.getXSSFWorkbook().getSheetIndex(sheet);
        if (!writeContext.getResolvedSheetRecord().containsKey(sheetIndex) || isFinal){
            HashBasedTable<Integer, String, CellAddress> singleReferenceData = writeContext.getSingleReferenceData();
            HashBasedTable<Integer, String, CellAddress> circleReferenceData = writeContext.getCircleReferenceData();
            HashBasedTable<Integer, String, CellAddress> calculateReferenceData = writeContext.getCalculateReferenceData();
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
            int singleReferenceDataSize = writeContext.getSingleReferenceData().size();
            int circleReferenceDataSize = writeContext.getCircleReferenceData().size();
            int calculateReferenceDataSize = writeContext.getCalculateReferenceData().size();
            writeContext.getResolvedSheetRecord().put(sheetIndex,true);
            debug(LOGGER, format("%s工作表索引[%s]解析模板完成，共解析到[%s]个占位符,引用占位符[%s]个,列表占位符[%s]个,计算占位符[%s]个",
                    isFinal? "[收尾阶段]":"",
                    sheetIndex,
                    singleReferenceDataSize + circleReferenceDataSize + calculateReferenceDataSize,
                    singleReferenceDataSize,
                    circleReferenceDataSize,
                    calculateReferenceDataSize
            ));
        }else{
            debug(LOGGER, format("工作表[%s]已被解析过，跳过本次解析",sheetIndex));
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
                LoggerHelper.debug(LOGGER, format("解析到占位符[%s]为合并单元格[%s]",cellAddress.getPlaceholder(),cellMerged.formatAsString()));
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
            for (Integer i : writeContext.getResolvedSheetRecord().keySet()) {
                this.resolveTemplate(getWorkbookSheet(i), true);
                this.gatherUnusedSingleReferenceDataAndFillDefault();
                this.gatherUnusedCircleReferenceDataAndFillDefault();
                this.setCalculateData(i);
            }
        }else{
            this.resolveTemplate(getWorkbookSheet(writeContext.getSwitchSheetIndex()),false);
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
        OutputStream outputStream = writeConfig.getOutputStream();
        if(outputStream != null){
            this.flush(true);
            workbook.write(writeConfig.getOutputStream());
            workbook.close();
            writeConfig.getOutputStream().close();
        }else{
            String message = "输出流为空,请指定输出流";
            debug(LOGGER,message);
            throw new AxolotlWriteException(message);
        }
    }

    @Override
    public void switchSheet(int sheetIndex) {
        super.switchSheet(sheetIndex);
        this.resolveTemplate(this.getWorkbookSheet(sheetIndex),false);
    }
}

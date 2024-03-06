package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.constant.TemplatePlaceholderPattern;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.*;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.axolotl.toolkit.tika.TikaShell;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.MapDifference;
import com.google.common.collect.Maps;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;

import java.io.File;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.debug;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.format;

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
        this.writeConfig = templateWriteConfig;
        TemplateWriteContext templateWriteContext = new TemplateWriteContext();
        super.writeContext = templateWriteContext;
        this.writeContext = templateWriteContext;
        this.writeContext.setSwitchSheetIndex(writeConfig.getSheetIndex());
        super.LOGGER = LOGGER;
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
            sheet = getWorkbookSheet(writeContext.getSwitchSheetIndex());
            // 只有第一次写入时解析模板占位符
            if (writeContext.isFirstBatch(writeContext.getSwitchSheetIndex())){
                // 解析模板占位符到上下文
                this.resolveTemplate(sheet,false);
            }
            // 写入Map映射
            this.writeSingleData(sheet, fixMapping,writeContext.getSingleReferenceData(),false);
            // 写入循环数据
            this.writeCircleData(sheet, dataList);
            axolotlWriteResult.setWrite(true);
            axolotlWriteResult.setMessage("写入完成");
        }else{
            String message = "非模板写入请使用AxolotlAutoExcelWriter.write()方法";
            if(writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.EXCEPTION_RETURN_RESULT)){
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
    private void writeSingleData(Sheet sheet,Map<String,?> singleMap,HashBasedTable<Integer, String, CellAddress> referenceData,boolean gatherUnusedStage){
        int sheetIndex = workbook.getXSSFWorkbook().getSheetIndex(sheet);
        Map<String, CellAddress> addressMapping = referenceData.row(sheetIndex);
        HashMap<String, Object> dataMapping = (singleMap != null ? new HashMap<>(singleMap) : new HashMap<>());
        this.injectCommonConstInfo(dataMapping,gatherUnusedStage);
        HashBasedTable<Integer, String, Boolean> alreadyUsedReferenceData = writeContext.getAlreadyUsedReferenceData();
        Map<String, Boolean> alreadyUsedDataMapping = alreadyUsedReferenceData.row(sheetIndex);
        for (String singleKey : addressMapping.keySet()) {
            CellAddress cellAddress = addressMapping.get(singleKey);
            String placeholder = cellAddress.getPlaceholder();
            if(alreadyUsedDataMapping.containsKey(placeholder)){continue;}
            Cell cell = sheet.getRow(cellAddress.getRowPosition()).getCell(cellAddress.getColumnPosition());
            if (dataMapping.containsKey(singleKey)){
                Object info = dataMapping.get(singleKey);
                if(info == null){
                    if(gatherUnusedStage){
                        debug(LOGGER, format("[收尾阶段]设置模板占位符[%s]为空值",placeholder));
                    }else {
                        debug(LOGGER, format("设置模板占位符[%s]为空值",placeholder));
                    }
                    cell.setBlank();
                }else{
                    debug(LOGGER, format("设置模板占位符[%s]值[%s]",placeholder,info));
                    cell.setCellValue(cellAddress.replacePlaceholder(info.toString()));
                }
                cellAddress.setWrittenRow(cell.getRowIndex());
                alreadyUsedDataMapping.put(cellAddress.getPlaceholder(),true);
            }else {
                debug(LOGGER, format("未找到模板占位符[%s]",placeholder));
            }
        }
    }

    /**
     * 未使用的单次占位符填充默认值
     */
    private void gatherUnusedSingleReferenceDataAndFillDefault() {
        if(writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.PLACEHOLDER_FILL_DEFAULT)){
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
        if(writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.PLACEHOLDER_FILL_DEFAULT)){
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
            unusedMap.put(singleKey,null);
        }
        return unusedMap;
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
        boolean dataNotEmpty = Validator.objNotNull(circleDataList);
        Map<String, CellAddress> circleReferenceData = this.writeContext.getCircleReferenceData().row(writeContext.getSwitchSheetIndex());
        if (dataNotEmpty){
            boolean isSimplePOJO;
            // 获取写入类字段数据
            Map<String,Integer> writeFieldNames = new HashMap<>();
            Object rowObjInstance = circleDataList.get(0);
            List<String> writeFieldNamesList;
            if (rowObjInstance instanceof Map){
                isSimplePOJO = false;
                Map<String, Object> rowObjInstanceMap = (Map<String, Object>) rowObjInstance;
                if (!rowObjInstanceMap.isEmpty()){
                    writeFieldNames = rowObjInstanceMap.keySet()
                            .stream()
                            .collect(Collectors.toMap(key -> key, key -> 1));
                }
            }else {
                isSimplePOJO = true;
                Class<?> instanceClass = rowObjInstance.getClass();
                for (String key : circleReferenceData.keySet()) {
                    Field field;
                    try {
                        field = instanceClass.getDeclaredField(key);
                    }catch(NoSuchFieldException noSuchFieldException){
                        field = null;
                    }
                    if (field != null){
                        writeFieldNames.put(key, 1);
                    }
                }
            }
            writeFieldNamesList = new ArrayList<>(writeFieldNames.keySet());
            LoggerHelper.debug(LOGGER,"本次写入字段为:%s",writeFieldNames.keySet());
            // 漂移写入特性
            int sheetIndex = writeContext.getSwitchSheetIndex();
            boolean initialWriting = writeContext.fieldsIsInitialWriting(sheetIndex,writeFieldNamesList);
            int startShiftRow = calculateStartShiftRow(circleReferenceData, writeFieldNames, initialWriting);
            boolean nonTemplateCellFill = writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.NON_TEMPLATE_CELL_FILL);
            if (initialWriting && nonTemplateCellFill){
                // 获取模板行次的非模板值的列
                int templateLineIdx = startShiftRow - 1;
                XSSFRow templateRow = sheet.getRow(templateLineIdx);
                int templateRowLastCellNum = templateRow.getLastCellNum();
                Map<Integer,String > templateColumnMap = circleReferenceData.values()
                        .stream().filter(cellAddress -> cellAddress.getRowPosition() == templateLineIdx)
                        .collect(Collectors.toMap(CellAddress::getColumnPosition, CellAddress::getPlaceholder));
                List<CellAddress> nonTemplateCellAddressList = new ArrayList<>();
                // 将非模板列存储
                for (int i = 0; i < templateRowLastCellNum; i++) {
                    if(!templateColumnMap.containsKey(i)){
                        XSSFCell cell = templateRow.getCell(i);
                        CellAddress nonTempalteCellAddress = new CellAddress(null, templateLineIdx, i, cell.getCellStyle());
                        nonTempalteCellAddress.set_nonTemplateCell(cell);
                        nonTempalteCellAddress.setMergeRegion(ExcelToolkit.isCellMerged(sheet, templateLineIdx, i));
                        nonTemplateCellAddressList.add(nonTempalteCellAddress);
                    }
                }
                // 存储非模板列
                writeContext.getSheetNonTemplateCells().put(sheetIndex,writeFieldNamesList,nonTemplateCellAddressList);
                debug(LOGGER,"获取模板行[%s]个非模板列",nonTemplateCellAddressList.size());
            }
            writeContext.addFieldRecords(sheetIndex,writeFieldNamesList,writeContext.getCurrentWrittenBatch().get(sheetIndex));
            if ((circleDataList.size() > 1 || (circleDataList.size() == 1 && initialWriting)) &&
                    writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SHIFT_WRITE_ROW)){
                // 最后一行大于起始行，则下移，否则为表底不下移
                if(sheet.getLastRowNum() >= startShiftRow){
                    int shiftRowNumber = initialWriting ? circleDataList.size() - 1 : circleDataList.size();
                    LoggerHelper.debug(LOGGER,"当前写入起始行次[%s],下移行次:[%s],",startShiftRow,shiftRowNumber);
                    sheet.shiftRows(startShiftRow, sheet.getLastRowNum(), shiftRowNumber, true,true);
                }
            }
            // 写入列表数据
            HashBasedTable<Integer, String, Boolean> alreadyUsedReferenceData = writeContext.getAlreadyUsedReferenceData();
            Map<String, Boolean> alreadyUsedDataMapping = alreadyUsedReferenceData.row(writeContext.getSwitchSheetIndex());
            Map<String, CellAddress> calculateReferenceData = this.writeContext.getCalculateReferenceData().row(writeContext.getSwitchSheetIndex());
            for (Object data : circleDataList) {
                debug(LOGGER,"[写入数据]"+data);
                for (Map.Entry<String, CellAddress> fieldMapping : circleReferenceData.entrySet()) {
                    String fieldMappingKey = fieldMapping.getKey();
                    CellAddress cellAddress = circleReferenceData.get(fieldMappingKey);
                    int rowPosition = cellAddress.getRowPosition();
                    if(writeFieldNames.containsKey(fieldMappingKey)){
                        Object value;
                        if (isSimplePOJO){
                            Field field = data.getClass().getDeclaredField(fieldMappingKey);
                            field.setAccessible(true);
                            value = field.get(data);
                        }else{
                            Map<String, Object> map = (Map<String, Object>) data;
                            value = map.get(fieldMappingKey);
                        }
                        Cell writableCell = ExcelToolkit.createOrCatchCell(sheet, rowPosition,
                                cellAddress.getColumnPosition(), cellAddress.getCellStyle());
                        // 空值时使用默认值填充
                        if (Validator.strIsBlank(value)){
                            String defaultValue = cellAddress.getDefaultValue();
                            if (defaultValue != null){
                                writableCell.setCellValue(cellAddress.replacePlaceholder(defaultValue));
                            }else{
                                if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.NULL_VALUE_WITH_TEMPLATE_FILL)){
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
                                this.writeContext.getCalculateReferenceData().put(sheetIndex,fieldMappingKey, calculateAddress);
                            }
                            // 暂时只适配String类型
                            writableCell.setCellValue(cellAddress.replacePlaceholder(value.toString()));
                        }
                    // 将未引用的占位符填充默认值
                    }else{
                        ExcelToolkit.cellAssignment(
                                sheet, rowPosition, cellAddress.getColumnPosition(),
                                cellAddress.getCellStyle(), cellAddress.getDefaultValue()
                        );
                    }
                    this.setMergeRegion(sheet,cellAddress,rowPosition);
                    cellAddress.setRowPosition(++rowPosition);
                    if (!alreadyUsedDataMapping.containsKey(cellAddress.getPlaceholder())){
                        alreadyUsedDataMapping.put(cellAddress.getPlaceholder(),true);
                    }
                }
                if (nonTemplateCellFill){
                    List<CellAddress> nonTemplateCells = writeContext.getSheetNonTemplateCells().get(sheetIndex, writeFieldNamesList);
                    if (nonTemplateCells != null) {
                        for (CellAddress nonTemplateCellAddress : nonTemplateCells){
                            int rowPosition = nonTemplateCellAddress.getRowPosition();
                            Cell writableCell = ExcelToolkit.createOrCatchCell(sheet, rowPosition, nonTemplateCellAddress.getColumnPosition(), null);
                            ExcelToolkit.cloneOldCell2NewCell(writableCell,nonTemplateCellAddress.get_nonTemplateCell());
                            this.setMergeRegion(sheet,nonTemplateCellAddress,rowPosition);
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
    private void setMergeRegion(Sheet sheet,CellAddress cellAddress,int rowPosition){
        if (cellAddress.isMergeCell() && !cellAddress.isInitializedWrite()){
            CellRangeAddress mergeRegion = cellAddress.getMergeRegion();
            mergeRegion.setFirstRow(rowPosition);
            mergeRegion.setLastRow(rowPosition);
            StyleHelper.renderMergeRegionStyle(sheet,mergeRegion,cellAddress.getCellStyle());
            sheet.addMergedRegion(mergeRegion);
        }
    }

    /**
     * 计算起始行
     * @param circleReferenceData 引用数据
     * @param writeFieldNames 写入字段
     * @param initialWriting 是否是第一次写入
     * @return 起始行
     */
    private static int calculateStartShiftRow(Map<String, CellAddress> circleReferenceData, Map<String, Integer> writeFieldNames, boolean initialWriting) {
        int maxRowPosition = Integer.MIN_VALUE;
        for (Map.Entry<String, CellAddress> addressEntry : circleReferenceData.entrySet()) {
            if (writeFieldNames.containsKey(addressEntry.getKey())){
                maxRowPosition = Math.max(maxRowPosition, addressEntry.getValue().getRowPosition());
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
                        if (cell != null && CellType.STRING.equals(cell.getCellType())){
                            String cellValue = cell.getStringCellValue();
                            CellAddress cellAddress = new CellAddress(cellValue,rowIdx, colIdx,cell.getCellStyle());
                            Boolean foundPlaceholder = findPlaceholderData(isFinal,singleReferenceData,
                                    TemplatePlaceholderPattern.SINGLE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cellAddress);
                            if (foundPlaceholder == null){
                                foundPlaceholder = findPlaceholderData(isFinal,circleReferenceData,
                                        TemplatePlaceholderPattern.CIRCLE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cellAddress);
                            }
                            if (foundPlaceholder == null) {
                                foundPlaceholder = findPlaceholderData(isFinal,calculateReferenceData,
                                        TemplatePlaceholderPattern.AGGREGATE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cellAddress);
                            }
                            if (foundPlaceholder != null && foundPlaceholder){
                                CellRangeAddress cellMerged = ExcelToolkit.isCellMerged(sheet, rowIdx, colIdx);
                                if (cellMerged != null){
                                    LoggerHelper.debug(LOGGER, format("解析到占位符[%s]为合并单元格[%s]",cellAddress.getPlaceholder(),cellMerged.formatAsString()));
                                    cellAddress.setMergeRegion(cellMerged);
                                }
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
     *
     * @param isFinal 是否是收尾阶段
     * @param referenceData 引用数据
     * @param pattern       模板匹配正则
     * @param sheetIndex    工作簿索引
     * @param cellAddress   单元格地址
     */
    private Boolean findPlaceholderData(boolean isFinal, HashBasedTable<Integer, String, CellAddress> referenceData,
                                        Pattern pattern, int sheetIndex, CellAddress cellAddress) {
        Matcher matcher = pattern.matcher(cellAddress.getCellValue());
        if (matcher.find()) {
            cellAddress.setPlaceholder(matcher.group());
            String name = matcher.group(1);
            String[] defaultSplitContent = name.split(":");
            name = defaultSplitContent[0];
            if (defaultSplitContent.length > 1){
                cellAddress.setDefaultValue(defaultSplitContent[1]);
            }
            boolean isCirclePattern = pattern.equals(TemplatePlaceholderPattern.CIRCLE_REFERENCE_TEMPLATE_PATTERN);
            if (isCirclePattern || pattern.equals(TemplatePlaceholderPattern.SINGLE_REFERENCE_TEMPLATE_PATTERN)){
                cellAddress.setPlaceholderType(isCirclePattern ? PlaceholderType.CIRCLE : PlaceholderType.MAPPING);
                referenceData.put(sheetIndex,name, cellAddress);
            }else if (pattern.equals(TemplatePlaceholderPattern.AGGREGATE_REFERENCE_TEMPLATE_PATTERN)){
                if (!isFinal){
                    cellAddress.setPlaceholderType(PlaceholderType.CALCULATE);
                    cellAddress.setCalculatedValue(BigDecimal.ZERO);
                }else {
                    Map<String, CellAddress> addressMap = referenceData.row(sheetIndex);
                    if (addressMap.containsKey(name)){
                        CellAddress originalAddress = addressMap.get(name);
                        originalAddress.setRowPosition(cellAddress.getRowPosition());
                        referenceData.put(sheetIndex,name, originalAddress);
                        return true;
                    }
                }
                referenceData.put(sheetIndex,name, cellAddress);
            }
            return true;
        }
        return null;
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
        this.flush(true);
        workbook.write(writeConfig.getOutputStream());
        workbook.close();
        writeConfig.getOutputStream().close();
    }

    @Override
    public void switchSheet(int sheetIndex) {
        super.switchSheet(sheetIndex);
        this.resolveTemplate(this.getWorkbookSheet(sheetIndex),false);
    }
}

package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.common.CommonMimeType;
import cn.toutatis.xvoid.axolotl.excel.writer.constant.TemplatePlaceholderPattern;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractInnerStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.CellAddress;
import cn.toutatis.xvoid.axolotl.excel.writer.support.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.support.PlaceholderType;
import cn.toutatis.xvoid.axolotl.excel.writer.support.WriteContext;
import cn.toutatis.xvoid.axolotl.manage.Progress;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.axolotl.toolkit.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.toolkit.tika.TikaShell;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import com.alibaba.fastjson.JSONObject;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.MapDifference;
import com.google.common.collect.Maps;
import lombok.SneakyThrows;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.slf4j.Logger;

import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.debug;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.format;

/**
 * 文档文件写入器
 * @author Toutatis_Gc
 */
public class AxolotlExcelWriter implements Closeable {

    /**
     * 日志记录器
     */
    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlExcelWriter.class);

    /**
     * 写入工作簿
     */
    private final SXSSFWorkbook workbook;

    /**
     * 写入上下文
     */
    private final WriteContext writeContext = new WriteContext();

    /**
     * 写入配置
     */
    private final WriterConfig writerConfig;

    /**
     * 主构造函数
     * @param writerConfig 写入配置
     */
    public AxolotlExcelWriter(WriterConfig writerConfig) {
        this.writerConfig = writerConfig;
        this.workbook = this.initWorkbook(null);
    }

    /**
     * 构造函数
     * 可以写入一个模板文件
     * @param templateFile 模板文件
     * @param writerConfig 写入配置
     */
    public AxolotlExcelWriter(File templateFile, WriterConfig writerConfig) {
        TikaShell.preCheckFileNormalThrowException(templateFile);
        this.writerConfig = writerConfig;
        this.workbook = this.initWorkbook(templateFile);
    }

    /**
     * 初始化工作簿
     * @param templateFile 模板文件
     * @return 工作簿
     */
    private SXSSFWorkbook initWorkbook(File templateFile) {
        SXSSFWorkbook workbook;
        // 读取模板文件内容
        if (templateFile != null){
            debug(LOGGER, format("正在使用模板文件[%s]作为写入模板",templateFile.getAbsolutePath()));
            TikaShell.preCheckFileNormalThrowException(templateFile);
            DetectResult detect = TikaShell.detect(templateFile, CommonMimeType.OOXML_EXCEL);
            if (!detect.isWantedMimeType()){
                throw new AxolotlWriteException("请使用xlsx文件作为写入模板");
            }
            this.writeContext.setTemplateFile(templateFile);
            try (FileInputStream fis = new FileInputStream(templateFile)){
                OPCPackage opcPackage = OPCPackage.open(fis);
                workbook = new SXSSFWorkbook(XSSFWorkbookFactory.createWorkbook(opcPackage));
            }catch (IOException | InvalidFormatException e){
                e.printStackTrace();
                throw new AxolotlWriteException(format("模板文件[%s]读取失败",templateFile.getAbsolutePath()));
            }
        }else {
            workbook = new SXSSFWorkbook();
        }
        return workbook;
    }



    @SneakyThrows
    public void write(){
        LoggerHelper.info(LOGGER, writeContext.getCurrentWrittenBatchAndIncrement());
        SXSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(writerConfig.getSheetIndex(),writerConfig.getSheetName());
        ExcelStyleRender styleRender = writerConfig.getStyleRender();
        if (styleRender instanceof AbstractInnerStyleRender innerStyleRender){
            innerStyleRender.setWriterConfig(writerConfig);
            innerStyleRender.renderHeader(sheet);
        }else {
            styleRender.renderHeader(sheet);
        }
        styleRender.renderData(sheet,writerConfig.getData());
        workbook.write(writerConfig.getOutputStream());
    }

    /**
     * 写入数据到模板
     * @param sheetIndex 工作簿索引
     * @param singleMap 固定值映射
     * @param circleDataList 循环数据
     */
    @SneakyThrows
    @SuppressWarnings("unchecked")
    public void writeToTemplate(int sheetIndex, Map<String,?> singleMap, List<?> circleDataList){
        LoggerHelper.info(LOGGER, writeContext.getCurrentWrittenBatchAndIncrement());
        XSSFSheet sheet;
        if (writeContext.isTemplateWrite()){
            sheet = workbook.getXSSFWorkbook().getSheetAt(sheetIndex);
            if (sheet != null){
                // 解析模板占位符到上下文
                this.resolveTemplate(sheet);
                // 写入Map映射
                this.writeSingleData(sheet,singleMap,false);
                boolean dataNotEmpty = Validator.objNotNull(circleDataList);
                String progressId = Progress.generateProgressId();
                Progress.init(progressId,dataNotEmpty? circleDataList.size() : 1);
                // 写入循环数据
                Map<String, CellAddress> circleReferenceData = this.writeContext.getCircleReferenceData().row(sheetIndex);
                if (dataNotEmpty){
                    boolean isSimplePOJO;
                    // 获取写入类字段数据
                    Map<String,Integer> writeFieldNames = new HashMap<>();
                    Object rowObjInstance = circleDataList.get(0);
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
                    LoggerHelper.debug(LOGGER,"本次写入字段为:%s",writeFieldNames.keySet());
                    // 漂移写入特性
                    if (circleDataList.size() > 1 && writerConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SHIFT_WRITE_ROW)){
                        int maxRowPosition = Integer.MIN_VALUE;
                        for (Map.Entry<String, CellAddress> addressEntry : circleReferenceData.entrySet()) {
                            if (writeFieldNames.containsKey(addressEntry.getKey())){
                                maxRowPosition = Math.max(maxRowPosition, addressEntry.getValue().getRowPosition());
                            }
                        }
                        sheet.shiftRows(maxRowPosition+1,sheet.getLastRowNum(),circleDataList.size()-1,true,true);
                    }
                    // 写入数据
                    for (Object data : circleDataList) {
                        for (Map.Entry<String, Integer> fieldMapping : writeFieldNames.entrySet()) {
                            CellAddress cellAddress = circleReferenceData.get(fieldMapping.getKey());
                            Object value;
                            if (isSimplePOJO){
                                Field field = data.getClass().getDeclaredField(fieldMapping.getKey());
                                field.setAccessible(true);
                                value = field.get(data);
                            }else{
                                Map<String, Object> map = (Map<String, Object>) data;
                                value = map.get(fieldMapping.getKey());
                            }
                            int rowPosition = cellAddress.getRowPosition();
                            boolean mergeCell = cellAddress.isMergeCell();
                            XSSFRow writableRow = sheet.getRow(rowPosition);
                            if (writableRow == null){
                                writableRow = sheet.createRow(rowPosition);
                            }
                            XSSFCell writableCell = writableRow.getCell(cellAddress.getColumnPosition());
                            if (writableCell == null){
                                writableCell = writableRow.createCell(cellAddress.getColumnPosition());
                            }
                            writableCell.setCellStyle(cellAddress.getCellStyle());
                            if (Validator.strIsBlank(value)){
                                writableCell.setBlank();
                            }else {
                                // 暂时只适配String类型
                                writableCell.setCellValue(cellAddress.replacePlaceholder(value.toString()));
                            }
                            if (mergeCell && !cellAddress.isInitializedWrite()){
                                CellRangeAddress mergeRegion = cellAddress.getMergeRegion();
                                mergeRegion.setFirstRow(rowPosition);
                                mergeRegion.setLastRow(rowPosition);
                                StyleHelper.renderMergeRegionStyle(sheet,mergeRegion,cellAddress.getCellStyle());
                                sheet.addMergedRegion(mergeRegion);
                            }
                            cellAddress.setRowPosition(++rowPosition);
                        }
                    }
                }
            }else{
                throw new AxolotlWriteException(format("工作表索引[%s]模板中不存在",sheetIndex));
            }
        }else{
            throw new AxolotlWriteException("非模板写入请使用write方法");
        }
    }

    private void writeSingleData(Sheet sheet,Map<String,?> singleMap,boolean gatherUnusedStage){
        HashBasedTable<Integer, String, CellAddress> singleReferenceData = this.writeContext.getSingleReferenceData();
        int sheetIndex = workbook.getXSSFWorkbook().getSheetIndex(sheet);
        Map<String, CellAddress> singleAddressMapping = singleReferenceData.row(sheetIndex);
        HashMap<String, Object> dataMapping = new HashMap<>(singleMap);
        this.injectCommonInfo(dataMapping,gatherUnusedStage);
        HashBasedTable<Integer, String, Boolean> alreadyUsedReferenceData = writeContext.getAlreadyUsedReferenceData();
        Map<String, Boolean> alreadyUsedDataMapping = alreadyUsedReferenceData.row(sheetIndex);
        for (String singleKey : singleAddressMapping.keySet()) {
            if(alreadyUsedDataMapping.containsKey(singleKey)){continue;}
            CellAddress cellAddress = singleAddressMapping.get(singleKey);
            Cell cell = sheet.getRow(cellAddress.getRowPosition()).getCell(cellAddress.getColumnPosition());
            if (dataMapping.containsKey(singleKey)){
                Object info = dataMapping.get(singleKey);
                if(info == null){
                    if(gatherUnusedStage){
                        debug(LOGGER, format("设置模板占位符[%s]为空值",singleKey));
                    }
                    cell.setBlank();
                }else{
                    debug(LOGGER, format("设置模板占位符[%s]值[%s]",singleKey,info));
                    cell.setCellValue(cellAddress.replacePlaceholder(info.toString()));
                }
                cellAddress.setWrittenRow(cell.getRowIndex());
                alreadyUsedDataMapping.put(singleKey,true);
            }else {
                debug(LOGGER, format("未找到模板占位符[%s]",singleKey));
            }
        }
    }

    /**
     * 未使用的占位符填充默认值
     */
    private void gatherUnusedSingleReferenceDataAndFillDefault() {
        if(writerConfig.getWritePolicyAsBoolean(ExcelWritePolicy.PLACEHOLDER_FILL_DEFAULT)){
            int sheetIndex = writerConfig.getSheetIndex();
            Sheet sheet = workbook.getXSSFWorkbook().getSheetAt(sheetIndex);
            this.resolveTemplate(sheet);
            Map<String, CellAddress> singleReferenceMapping =  writeContext.getSingleReferenceData().row(sheetIndex);
            Map<String, Boolean> alreadyUsedDataMapping =  writeContext.getAlreadyUsedReferenceData().row(sheetIndex);
            MapDifference<String, Object> difference = Maps.difference(singleReferenceMapping, alreadyUsedDataMapping);
            Map<String, Object> onlyOnLeft = difference.entriesOnlyOnLeft();
            HashMap<String, Object> unusedMap = new HashMap<>();
            for (String singleKey : onlyOnLeft.keySet()) {
                unusedMap.put(singleKey,null);
            }
            this.writeSingleData(sheet,unusedMap,true);

        }
    }

    private void gatherUnusedCircleReferenceDataAndFillDefault() {
        int sheetIndex = writerConfig.getSheetIndex();
        Map<String, CellAddress> singleReferenceMapping =  writeContext.getCircleReferenceData().row(sheetIndex);
        System.err.println(singleReferenceMapping);
        // TODO 采集未使用循环实体
    }

    /**
     * 解析数据实体类型
     * @param data 数据集合
     */
    private void writeCircleData(Sheet sheet,List<Object> data,Map<String, CellAddress> circleReferenceData){
        // TODO 获取实体字段名
        if (data == null || data.isEmpty()){
            return;
        }
        Object dataTmp = data.get(0);
        Class<?> dataClass = dataTmp.getClass();
        if (dataTmp instanceof Map<?,?>){

        }
    }

    @SuppressWarnings({"rawtypes","unchecked"})
    private void injectCommonInfo(Map singleMap,boolean gatherUnusedStage){
        if(!gatherUnusedStage){
            if (singleMap == null){
                singleMap = new HashMap<>();
            }
            singleMap.put("AXOLOTL_CREATE_TIME", Time.getCurrentTime());
        }
    }

    /**
     * 解析模板
     * @param sheet 工作表
     */
    private void resolveTemplate(Sheet sheet){
        int lastRowNum = sheet.getLastRowNum();
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (int rowIdx = 0; rowIdx <= lastRowNum; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row != null){
                short lastCellNum = row.getLastCellNum();
                for (int colIdx = 0; colIdx < lastCellNum; colIdx++) {
                    Cell cell = row.getCell(colIdx);
                    if (cell != null && CellType.STRING.equals(cell.getCellType())){
                        String cellValue = cell.getStringCellValue();
                        int sheetIndex = workbook.getXSSFWorkbook().getSheetIndex(sheet);
                        CellAddress cellAddress = new CellAddress(cellValue,rowIdx, colIdx,cell.getCellStyle());
                        Boolean foundPlaceholder = findPlaceholderData(
                                writeContext.getSingleReferenceData(),
                                TemplatePlaceholderPattern.SINGLE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cellAddress
                        );

                        if (foundPlaceholder == null){
                            foundPlaceholder = findPlaceholderData(
                                    writeContext.getCircleReferenceData(),
                                    TemplatePlaceholderPattern.CIRCLE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cellAddress
                            );
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
        debug(LOGGER, format("解析模板完成，共解析到%s个占位符",writeContext.getSingleReferenceData().size() + writeContext.getCircleReferenceData().size()));
    }

    /**
     * 解析模板值到变量
     * @param referenceData 引用数据
     * @param pattern 模板匹配正则
     * @param sheetIndex 工作簿索引
     * @param cellAddress 单元格地址
     */
    private Boolean findPlaceholderData(HashBasedTable<Integer, String, CellAddress> referenceData, Pattern pattern, int sheetIndex, CellAddress cellAddress) {
        Matcher matcher = pattern.matcher(cellAddress.getCellValue());
        if (matcher.find()) {
            cellAddress.setPlaceholder(matcher.group());
            if (pattern.equals(TemplatePlaceholderPattern.CIRCLE_REFERENCE_TEMPLATE_PATTERN)){
                cellAddress.setPlaceholderType(PlaceholderType.CIRCLE);
            }else {
                cellAddress.setPlaceholderType(PlaceholderType.MAPPING);
            }
            referenceData.put(sheetIndex, matcher.group(1), cellAddress);
            return true;
        }
        return null;
    }

    public static void main(String[] args) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\" + IdUtil.randomUUID() + ".xlsx"));

        WriterConfig writerConfig = new WriterConfig();
        writerConfig.setTitle("测试生成表标题");
        ArrayList<String> columnNames = new ArrayList<>();
        columnNames.add("名称");
        columnNames.add("姓名");
        columnNames.add("性别");
        columnNames.add("身份证号");
        columnNames.add("地址");
        writerConfig.setColumnNames(columnNames);
        ArrayList<JSONObject> data = new ArrayList<>();
        for (int i = 0; i < 50; i++) {
            JSONObject json = new JSONObject(true);
            json.put("name", "name" + i);
            json.put("age", i);
            json.put("sex", i % 2 == 0? "男" : "女");
            json.put("card", 555444114);
            json.put("address", null);
            data.add(json);
        }
        writerConfig.setData(data);
        writerConfig.setOutputStream(fileOutputStream);
        AxolotlExcelWriter writer = new AxolotlExcelWriter(writerConfig);
        // TODO 写入
        writer.write();
        writer.close();

    }

    /**
     * 写入器刷新内容
     * 进入写入剩余内容进入关闭流前的收尾工作
     * @param isFinal 是否是最终刷新，关闭写入前的最后一次刷新
     */
    public void flush(boolean isFinal) {
        // 采集未映射数据
        // 暂时没有false的情况
        if (isFinal || true){
            this.writeContext.getSingleReferenceData().row(this.writerConfig.getSheetIndex()).clear();
            this.gatherUnusedSingleReferenceDataAndFillDefault();
            this.gatherUnusedCircleReferenceDataAndFillDefault();
        }
    }


    public void flush() {
        // 采集未映射数据
        this.flush(false);
    }

    /**
     * 关闭工作簿所对应输出流
     * @throws IOException IO异常
     */
    @Override
    public void close() throws IOException {
        LoggerHelper.debug(LOGGER, "工作薄写入进入关闭阶段");
        this.flush(true);
        workbook.write(writerConfig.getOutputStream());
        workbook.close();
        writerConfig.getOutputStream().close();
    }

}

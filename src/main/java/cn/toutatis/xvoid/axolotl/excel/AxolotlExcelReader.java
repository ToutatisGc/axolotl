package cn.toutatis.xvoid.axolotl.excel;

import cn.toutatis.xvoid.axolotl.excel.constant.AxolotlDefaultConfig;
import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.constant.ReadExcelFeature;
import cn.toutatis.xvoid.axolotl.excel.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.support.exceptions.AxolotlReadException;
import cn.toutatis.xvoid.axolotl.excel.support.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.excel.support.tika.TikaShell;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import lombok.Getter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.RecordFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class AxolotlExcelReader<T> {

    /**
     * 日志
     */
    private final Logger LOGGER  = LoggerToolkit.getLogger(AxolotlExcelReader.class);

    /**
     * 工作簿元信息
     */
    @Getter
    private WorkBookContext workBookContext;

    /**
     * 构造文件读取器
     */
    public AxolotlExcelReader(File excelFile) {
        this(excelFile,true);
    }

    @SuppressWarnings("unchecked")
    public AxolotlExcelReader(File excelFile, boolean withDefaultConfig) {
        this(excelFile, (Class<T>) Object.class,withDefaultConfig);
    }

    public AxolotlExcelReader(File excelFile, Class<T> clazz) {
        this(excelFile,clazz,true);
    }

    /**
     * [ROOT]
     * 构造文件读取器
     * @param excelFile Excel工作簿文件
     * @param withDefaultConfig 是否使用默认配置
     */
    public AxolotlExcelReader(File excelFile, Class<T> clazz, boolean withDefaultConfig) {
        if (clazz == null){
            throw new IllegalArgumentException("读取的类型对象不能为空");
        }
        this.detectFileAndInitWorkbook(excelFile);
        this.workBookContext.setDirectReadClass(clazz);
        this.workBookContext.setUseDefaultReaderConfig(withDefaultConfig);
    }

    /**
     * 初始化读取Excel文件
     * 1.初始化加载文件先判断文件是否正常并且是需要的格式
     * 2.将文件加载到POI工作簿中
     * @param excelFile Excel工作簿文件
     */
    private void detectFileAndInitWorkbook(File excelFile) {
        // 检查文件格式是否为XLSX
        DetectResult detectResult = TikaShell.detect(excelFile, TikaShell.OOXML_EXCEL,false);
        if (!detectResult.isDetect()){
            DetectResult.FileStatus currentFileStatus = detectResult.getCurrentFileStatus();
            if (currentFileStatus == DetectResult.FileStatus.FILE_MIME_TYPE_PROBLEM ||
                    currentFileStatus == DetectResult.FileStatus.FILE_SUFFIX_PROBLEM
            ){
                //如果是因为后缀不匹配或媒体类型不匹配导致识别不通过 换XLS格式再次识别
                detectResult = TikaShell.detect(excelFile, TikaShell.MS_EXCEL,false);
            }else {
                //如果是预检查不通过抛出异常
                detectResult.throwException();
            }
        }
        // 检查文件是否正常并且是需要的类型，否则抛出异常
        if (detectResult.isDetect()){
            workBookContext = new WorkBookContext(excelFile,detectResult);
        }else{
            detectResult.throwException();
        }
        // 读取文件加载到元信息
        try(FileInputStream fis = new FileInputStream(workBookContext.getFile())){
            // 校验文件大小
            Workbook workbook;
            if (detectResult.getCatchMimeType() == TikaShell.OOXML_EXCEL && excelFile.length() > AxolotlDefaultConfig.XVOID_DEFAULT_MAX_FILE_SIZE){
                this.workBookContext.setEventDriven(true);
                OPCPackage opcPackage = OPCPackage.open(fis);
                workbook = XSSFWorkbookFactory.createWorkbook(opcPackage);
                opcPackage.close();
            }else {
                workbook = WorkbookFactory.create(fis);
            }
            workBookContext.setWorkbook(workbook);
        } catch (IOException | RecordFormatException | InvalidFormatException e) {
            LOGGER.error("加载文件失败",e);
            throw new AxolotlReadException(e.getMessage());
        }
    }

    /**
     * [ROOT]
     * 读取Excel数据
     * @param readerConfig 读取配置
     * @return 读取数据
     * @param <RT> 读取的类型泛型
     */
    public <RT> List<RT> readSheetData(ReaderConfig<RT> readerConfig) {
        // 检查并修正配置文件
        this.preCheckAndFixReadConfig(readerConfig);
        List<RT> readResult = new ArrayList<>();
        Sheet sheet = workBookContext.getWorkbook().getSheetAt(readerConfig.getSheetIndex());
        if (sheet == null){
            String msg = "读取的sheet不存在[%s]".formatted(readerConfig.getSheetIndex());
            if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.IGNORE_EMPTY_SHEET_ERROR)){
                LOGGER.warn(msg+"将返回空数据");
                return readResult;
            }else{
                LOGGER.error(msg);
                throw new AxolotlReadException(msg);
            }
        }
        this.readSheetData(sheet,readerConfig,readResult);
        return readResult;
    }

    /**
     * 读取表中每一行的数据
     */
    private <RT> void readSheetData(Sheet sheet,ReaderConfig<RT> readerConfig,List<RT> list){
        int startIndex = readerConfig.getStartIndex();
        int endIndex = readerConfig.getEndIndex();
        if (startIndex == 0){
            int initialRowPositionOffset = readerConfig.getInitialRowPositionOffset();
            if (initialRowPositionOffset > 0){
                LOGGER.debug("跳过前{}行",initialRowPositionOffset);
                startIndex = startIndex + initialRowPositionOffset;
                endIndex = endIndex + initialRowPositionOffset;
            }
        }
        for (int i = startIndex; i < endIndex; i++) {
            RT instance = this.readRow(sheet, i, readerConfig);
            if (instance!= null){list.add(instance);}
        }
    }

    /**
     * [ROOT]
     * 读取行信息到对象
     * @param sheet 表
     * @param rowNumber 行号
     * @param readerConfig 读取配置
     * @param <RT> 转换类型
     */
    private <RT> RT readRow(Sheet sheet,int rowNumber,ReaderConfig<RT> readerConfig){
        RT instance = readerConfig.getCastClassInstance();
        Row row = sheet.getRow(rowNumber);
        if (row == null){
            if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.INCLUDE_EMPTY_ROW)){
                return instance;
            }else{
                return null;
            }
        }
        this.convertCellToInstance(row,instance,readerConfig);
        return instance;
    }

    @SuppressWarnings({"unchecked","rawtypes"})
    private <RT> void convertCellToInstance(Row row,RT instance,ReaderConfig<RT> readerConfig){
        if (instance instanceof Map mapInstance){
            this.row2MapInstance(mapInstance, row,readerConfig);
        }else{
            this.row2SimplePOJO(instance,row,readerConfig);
        }
    }

    /**
     * 填充单元格数据到Java对象
     * @param instance 对象实例
     * @param row 行
     * @param readerConfig 读取配置
     * @param <RT> 读取类型
     */
    private <RT> void row2SimplePOJO(RT instance,Row row,ReaderConfig<RT> readerConfig){
        workBookContext.setCurrentReadRowIndex(row.getRowNum());
        List<EntityCellMappingInfo<?>> indexMappingInfos = readerConfig.getIndexMappingInfos();
        Map<String, EntityCellMappingInfo<?>> positionMappingInfos = readerConfig.getPositionMappingInfos();
        for (EntityCellMappingInfo<?> indexMappingInfo : indexMappingInfos) {

        }
        row.cellIterator().forEachRemaining(cell -> {
            workBookContext.setCurrentReadColumnIndex(cell.getColumnIndex());
        });
    }

    /**
     * [ROOT]
     * 填充单元格数据到map
     */
    private <RT> void row2MapInstance(Map<String,Object> instance, Row row,ReaderConfig<RT> readerConfig){
        workBookContext.setCurrentReadRowIndex(row.getRowNum());
        row.cellIterator().forEachRemaining(cell -> {
            workBookContext.setCurrentReadColumnIndex(cell.getColumnIndex());
            int idx = cell.getColumnIndex() + 1;
            String key = "CELL_" + idx;
            //FIXME
//            instance.put(key,getCellValue(cell, null).getCellValue());
            if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.USE_MAP_DEBUG)){
                instance.put("CELL_TYPE_"+idx,cell.getCellType());
                if (cell.getCellType() == CellType.NUMERIC){
                    if (DateUtil.isCellDateFormatted(cell)){
                        instance.put("CELL_TYPE_"+idx,cell.getCellType());
                        instance.put("CELL_DATE_"+idx, Time.regexTime(cell.getDateCellValue()));
                    }else{
                        instance.put("CELL_TYPE_"+idx,cell.getCellType());
                    }
                }else {
                    instance.put("CELL_TYPE_"+idx,cell.getCellType());
                }
            }
        });
    }



    /**
     * [ROOT]
     * 预校验读取配置是否正常
     * 不正常的数据将被修正
     * @param readerConfig 读取配置
     */
    private void preCheckAndFixReadConfig(ReaderConfig<?> readerConfig) {
        //检查部分
        if (readerConfig == null){
            String msg = "读取配置不能为空";
            LOGGER.error(msg);
            throw new AxolotlReadException(msg);
        }
        int sheetIndex = readerConfig.getSheetIndex();
        if (sheetIndex < 0){
            throw new AxolotlReadException("读取的sheetIndex不能小于0");
        }
        Class<?> castClass = readerConfig.getCastClass();
        if (castClass == null){
            throw new AxolotlReadException("读取的类型对象不能为空");
        }
        if (readerConfig.getStartIndex() < 0){
            throw new AxolotlReadException("读取起始行不得小于0");
        }
        if (readerConfig.getEndIndex() < 0){
            throw new AxolotlReadException("读取结束行不得小于0");
        }
        //修正部分
        if (readerConfig.getInitialRowPositionOffset() < 0){
            LOGGER.warn("读取的初始行偏移量不能小于0，将被修正为0");
            readerConfig.setInitialRowPositionOffset(0);
        }

    }


    /**
     * [ROOT]
     * 计算单元格公式为结果
     * @param cell 单元格
     * @return 计算结果
     */
    private CellGetInfo getFormulaCellValue(Cell cell) {
        // 从元数据中获取计算计算器
        CellValue evaluated = workBookContext.getFormulaEvaluator().evaluate(cell);
        // 将单元格为公式的单元格值转换为计算结果
        Object value = switch (evaluated.getCellType()) {
            case STRING -> evaluated.getStringValue();
            case NUMERIC -> evaluated.getNumberValue();
            case BOOLEAN -> evaluated.getBooleanValue();
            default -> {
                String msg = String.format("未知的公式单元格类型位置:[%d,%d],单元格类型:[%s],单元格值:[%s]",
                        cell.getRowIndex(), cell.getColumnIndex(), evaluated.getCellType(), evaluated);
                LOGGER.error(msg);
                throw new AxolotlReadException(msg);
            }
        };
        CellGetInfo cellGetInfo = new CellGetInfo(true, value);
        cellGetInfo.setCellType(cell.getCellType());
        return cellGetInfo;
    }

}

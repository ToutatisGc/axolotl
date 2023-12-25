package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.constant.ReadExcelFeature;
import cn.toutatis.xvoid.axolotl.support.DetectResult;
import cn.toutatis.xvoid.axolotl.support.TikaShell;
import cn.toutatis.xvoid.axolotl.support.WorkBookMetaInfo;
import cn.toutatis.xvoid.axolotl.support.WorkBookReaderConfig;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
public class GracefulExcelReader {

    private final Logger LOGGER  = LoggerToolkit.getLogger(GracefulExcelReader.class);

    @Getter
    private WorkBookMetaInfo workBookMetaInfo;

    @Getter
    @SuppressWarnings("rawtypes")
    private final WorkBookReaderConfig workBookReaderConfig;

    public GracefulExcelReader(File excelFile) {
        this(excelFile,true);
    }

    /**
     * 构造文件读取器
     * @param excelFile Excel工作簿文件
     * @param withDefaultConfig 是否使用默认配置
     */
    public GracefulExcelReader(File excelFile,boolean withDefaultConfig) {
        this.initWorkbook(excelFile);
        workBookReaderConfig = new WorkBookReaderConfig<>(withDefaultConfig);
    }

    @SuppressWarnings("unchecked")
    public <T> List<T> readSheetData(int sheetIndex,Class<T> clazz) {
        if (clazz == null || clazz == Object.class){
            throw new IllegalArgumentException("读取的类型对象不能为空");
        }
        if (sheetIndex < 0){
            boolean ignoreEmptySheetError = workBookReaderConfig.getReadFeatureAsBoolean(ReadExcelFeature.IGNORE_EMPTY_SHEET_ERROR);
            if (ignoreEmptySheetError){
                return null;
            }else{
                String msg = workBookReaderConfig.getSheetName() != null ? "表名[" + workBookReaderConfig.getSheetName() + "]不存在" : "表索引[" + sheetIndex + "]不存在";
                LOGGER.error(msg);
                throw new IllegalArgumentException(msg);
            }
        }
        workBookReaderConfig.setCastClass(clazz);
        workBookReaderConfig.setSheetIndex(sheetIndex);
        Sheet sheetAt = workBookMetaInfo.getWorkbook().getSheetAt(sheetIndex);
        int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
        int lastRowNum = sheetAt.getLastRowNum();
        return (List<T>) loadData(0,lastRowNum);
    }

    public <T> List<T> readSheetData(String sheetName, Class<T> clazz) {
        if (Validator.strIsBlank(sheetName)){throw new IllegalArgumentException("表名不能为空");}
        workBookReaderConfig.setSheetName(sheetName);
        int sheetIndex = this.workBookMetaInfo.getWorkbook().getSheetIndex(sheetName);
        return readSheetData(sheetIndex,clazz);
    }

    /**
     * 读取Excel文件
     * @param excelFile Excel工作簿文件
     */
    private void initWorkbook(File excelFile) {
        // 检查文件是否正常
        TikaShell.preCheckFileNormalThrowException(excelFile);
        DetectResult detectResult = TikaShell.detect(excelFile, TikaShell.OOXML_EXCEL,true);
        if (!detectResult.isDetect()){
            detectResult = TikaShell.detect(excelFile, TikaShell.MS_EXCEL,true);
        }
        if (detectResult.isDetect() && detectResult.isWantedMimeType()){
            workBookMetaInfo = new WorkBookMetaInfo(excelFile,detectResult);
        }else{
            detectResult.throwException();
        }
        // 读取文件加载到元信息
        try(FileInputStream fis = new FileInputStream(workBookMetaInfo.getFile())){
            Workbook workbook;
            if (workBookMetaInfo.getMimeType().equals(TikaShell.OOXML_EXCEL)){
                workbook = new XSSFWorkbook(fis);
            }else{
                workbook = new HSSFWorkbook(fis);
            }
            workBookMetaInfo.setWorkbook(workbook);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 加载数据
     * @param start 开始行
     * @param end 结束行
     */
    private List<Object> loadData(int start,int end){
        // 读取指定sheet
        Sheet sheet = workBookMetaInfo.getWorkbook().getSheetAt(workBookReaderConfig.getSheetIndex());
        int sheetIndex = workBookReaderConfig.getSheetIndex();
        List<Row> rowList;
        if (workBookMetaInfo.isSheetDataEmpty(sheetIndex)){
            ArrayList<Row> tmp = new ArrayList<>();
            if (workBookReaderConfig.getReadFeatureAsBoolean(ReadExcelFeature.INCLUDE_EMPTY_ROW)){
                for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    LOGGER.debug(i+":"+row);
                    tmp.add(row);
                }
            }else{
                sheet.rowIterator().forEachRemaining(tmp::add);
            }
            workBookMetaInfo.setSheetData(sheetIndex,tmp);
            rowList = tmp;
        }else {
            rowList = workBookMetaInfo.getSheetData(sheetIndex);
        }
        List<Object> dataList = new ArrayList<>();
        for (int idx = start; idx < end+1; idx++) {
            Object castClassInstance = workBookReaderConfig.getCastClassInstance();
            Row row = rowList.get(idx);
            if (row == null){
                dataList.add(castClassInstance);
                continue;
            }
            row.cellIterator().forEachRemaining(cell -> putCellToInstance(castClassInstance,cell));
            dataList.add(castClassInstance);
        }
        return dataList;
    }

    @SuppressWarnings({"unchecked","rawtypes"})
    private void putCellToInstance(Object instance, Cell cell){
        if (instance instanceof Map info){
            int idx = cell.getColumnIndex() + 1;
            String key = "CELL_" + idx;
            info.put(key,getCellValue(cell));
            if (workBookReaderConfig.getReadFeatureAsBoolean(ReadExcelFeature.USE_MAP_DEBUG)){
                info.put("CELL_TYPE_"+idx,cell.getCellType());
                if (cell.getCellType() == CellType.NUMERIC){
                    if (DateUtil.isCellDateFormatted(cell)){
                        info.put("CELL_TYPE_"+idx,"DATE");
                        info.put("CELL_DATE_"+idx, Time.regexTime(cell.getDateCellValue()));
                    }else{
                        info.put("CELL_TYPE_"+idx,cell.getCellType());
                    }
                }else {
                    info.put("CELL_TYPE_"+idx,cell.getCellType());
                }
            }
        }else{
            //TODO 一般POJO类型填充

        }
    }

    /**
     * 获取单元格值
     * @param cell 单元格
     * @return 单元格值
     */
    private Object getCellValue(Cell cell){
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                if (workBookReaderConfig.getReadFeatureAsBoolean(ReadExcelFeature.CAST_NUMBER_TO_DATE) && DateUtil.isCellDateFormatted(cell)){
                    yield cell.getDateCellValue();
                }else{
                    yield cell.getNumericCellValue();
                }
            }
            case BOOLEAN -> cell.getBooleanCellValue();
            case FORMULA -> getFormulaCellValue(cell);
            default -> {
                LOGGER.error("未知的单元格类型:{},{}",cell.getCellType(),cell);
                yield null;
            }
        };
    }

    /**
     * 计算公式
     * @param cell 单元格
     * @return 计算结果
     */
    private Object getFormulaCellValue(Cell cell){
        CellValue evaluated = workBookMetaInfo.getFormulaEvaluator().evaluate(cell);
        return switch (evaluated.getCellType()) {
            case STRING -> evaluated.getStringValue();
            case NUMERIC -> evaluated.getNumberValue();
            case BOOLEAN -> evaluated.getBooleanValue();
            default -> {
                LOGGER.error("未知的单元格类型:{},{}",evaluated.getCellType(), evaluated);
                yield null;
            }
        };
    }

    private void castAnyToString(Sheet sheet){
//        Sheet sheetAt = workbook.getSheetAt(0);
//        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
//        sheetAt.rowIterator().forEachRemaining(row -> {
//            row.cellIterator().forEachRemaining(cell -> {
//                switch (cell.getCellType()){
//                    case STRING:
//                        LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getStringCellValue());
//                        break;
//                    case BOOLEAN:
//                        LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getBooleanCellValue());
//                        break;
//                    case NUMERIC:
//                        LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getNumericCellValue());
//                        break;
//                    case FORMULA:
////                            LOGGER.info(cell.getCellFormula());
//                        CellValue evaluate = evaluator.evaluate(cell);
//                        System.err.println(evaluate.getNumberValue());
//                        break;
//                    default:
//                        LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.toString());
////                        case FORMULA:
////                            LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getCellFormula());
//                }
////                    LOGGER.info("row:{},cell:{},cellType:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getCellType());
//            });
//        });
    }

    public void reuse(){
        // TODO 用新的配置复用该文件对象
    }
}

package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.constant.ReadExcelFeature;
import cn.toutatis.xvoid.axolotl.support.DetectResult;
import cn.toutatis.xvoid.axolotl.support.TikaShell;
import cn.toutatis.xvoid.axolotl.support.WorkBookMetaInfo;
import cn.toutatis.xvoid.axolotl.support.WorkBookReaderConfig;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Getter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class GracefulExcelReader {

    private final Logger LOGGER  = LoggerToolkit.getLogger(GracefulExcelReader.class);

    @Getter
    private WorkBookMetaInfo workBookMetaInfo;

    @Getter
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
        workBookReaderConfig = new WorkBookReaderConfig(withDefaultConfig);
    }

    public <T> List<T> readSheetData(int sheetIndex,Class<T> clazz) {
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
        workBookReaderConfig.setSheetIndex(sheetIndex);
        return null;
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

    private void loadData(){
        Sheet sheet = workBookMetaInfo.getWorkbook().getSheetAt(workBookReaderConfig.getSheetIndex());

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
}

package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.support.DetectResult;
import cn.toutatis.xvoid.axolotl.support.TikaShell;
import cn.toutatis.xvoid.axolotl.support.WorkBookMetaInfo;
import cn.toutatis.xvoid.axolotl.support.WorkBookReaderConfig;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class GracefulExcelReader {

    private final Logger LOGGER  = LoggerToolkit.getLogger(GracefulExcelReader.class);

    @Getter
    private WorkBookMetaInfo workBookMetaInfo;

    @Setter
    @Getter
    private WorkBookReaderConfig workBookReaderConfig;


    /**
     * 构造文件读取器
     * @param excelFile Excel工作簿文件
     */
    public GracefulExcelReader(File excelFile) {
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
    }

    public void readData() {
        try(FileInputStream fis = new FileInputStream(workBookMetaInfo.getFile())){
            Workbook workbook;
            if (workBookMetaInfo.getMimeType().equals(TikaShell.OOXML_EXCEL)){
                workbook = new XSSFWorkbook(fis);
            }else{
                workbook = new HSSFWorkbook(fis);
            }
            HSSFSheet sheetAt = (HSSFSheet) workbook.getSheetAt(0);
//            List<Record> records = sheetAt.getWorkbook().getWorkbook().getRecords();
//            for (Record record : records) {
//                System.err.println(record.toString());
//            }
            sheetAt.rowIterator().forEachRemaining(row -> {
                row.cellIterator().forEachRemaining(cell -> {
                    switch (cell.getCellType()){
                        case STRING:
                            LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getStringCellValue());
                            break;
                        case BOOLEAN:
                            LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getBooleanCellValue());
                            break;
                        case NUMERIC:
                            LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getNumericCellValue());
                            break;
                        case FORMULA:
//                            LOGGER.info(cell.getCellFormula());
                            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
                            CellValue evaluate = formulaEvaluator.evaluate(cell);
                            System.err.println(evaluate.toString());
//                            cell.
                            break;
                        default:
                            LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.toString());
//                        case FORMULA:
//                            LOGGER.info("row:{},cell:{},type:{},value:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getCellFormula());
                    }
//                    LOGGER.info("row:{},cell:{},cellType:{}",row.getRowNum()+1,cell.getColumnIndex()+1,cell.getCellType(),cell.getCellType());
                });
            });
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}

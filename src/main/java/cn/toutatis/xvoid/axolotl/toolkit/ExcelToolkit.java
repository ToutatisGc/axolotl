package cn.toutatis.xvoid.axolotl.toolkit;

import cn.toutatis.xvoid.axolotl.Meta;
import cn.toutatis.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkitKt;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;

import java.util.Iterator;

/**
 * Excel工具类
 * @author Toutatis_Gc
 */
public class ExcelToolkit {

    private static final Logger LOGGER = LoggerToolkit.getLogger("Axolotl");

    /**
     * 获取当前读取到的行和列号的可读字符串
     * @return 当前读取到的行和列号的可读字符串
     */
    public static String getHumanReadablePosition(int rowIndex, int columnIndex) {
        char i = (char) ( 'A' + columnIndex);
        return String.format("%s", i+(String.format("%d",rowIndex + 1)));
    }

    /**
     * 判断当前行是否是空行
     * @param row 当前行
     * @return 当前行是否是空行
     */
    public static boolean blankRowCheck(Row row,int rangeStart,int rangeEnd){
        if (row == null){
            return true;
        }
        int isAllBlank = 0;
        short lastCellNum = rangeEnd < 0 ? row.getLastCellNum() : (short) rangeEnd;
        if(rangeStart > lastCellNum){
            throw new IllegalArgumentException("读取列起始位置必须大于结束位置");
        }
        for (int i = rangeStart; i < lastCellNum; i++) {
            Cell cell = row.getCell(i);
            if (cell == null || cell.getCellType() == CellType.BLANK){
                isAllBlank++;
            }else {
                return false;
            }
        }
        boolean blankRow = isAllBlank == lastCellNum;
        LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME, String.format("行[%s]数据为空",row.getRowNum()));
        return blankRow;
    }

    /**
     * 判断当前行是否是空行
     * @param row 当前行
     * @param readerConfig 读取配置
     * @return 当前行是否是空行
     */
    public static boolean blankRowCheck(Row row, ReaderConfig<?> readerConfig){
        int[] sheetColumnEffectiveRange = readerConfig.getSheetColumnEffectiveRange();
        return blankRowCheck(row, sheetColumnEffectiveRange[0],sheetColumnEffectiveRange[1]);
    }

    /**
     * 判断当前行不是空行
     * @param row 当前行
     * @return 当前行是否不是空行
     */
    public static boolean notBlankRowCheck(Row row, int rangeStart, int rangeEnd){
        return !blankRowCheck(row,rangeStart,rangeEnd);
    }

    /**
     * 判断当前行不是空行
     * @param row 当前行
     * @return 当前行是否不是空行
     */
    public static boolean notBlankRowCheck(Row row){
        return !blankRowCheck(row,0,-1);
    }

    public static void cloneOldWorkbook2NewWorkbook(Workbook newWorkbook, Workbook oldWorkBook){
        if (oldWorkBook == null || newWorkbook == null){return;}
        Iterator<Sheet> sheetIterator = oldWorkBook.sheetIterator();
        while (sheetIterator.hasNext()) {
            Sheet tmpSheet = sheetIterator.next();
            Sheet sxssfSheet = newWorkbook.createSheet(tmpSheet.getSheetName());
            cloneOldSheet2NewSheet(sxssfSheet, tmpSheet);
        }
    }

    public static void cloneOldSheet2NewSheet(Sheet newSheet, Sheet oldSheet){
        for (Row tmpRow : oldSheet) {
            if (tmpRow == null){continue;}
            Row sxssfRow = newSheet.createRow(tmpRow.getRowNum());
            cloneOldRow2NewRow(sxssfRow, tmpRow);
        }
    }

    public static void cloneOldRow2NewRow(Row newRow, Row oldRow){
        Iterator<Cell> cellIterator = oldRow.cellIterator();
        while (cellIterator.hasNext()) {
            Cell tmpCell = cellIterator.next();
            if (tmpCell == null){continue;}
            Cell newCell = newRow.createCell(tmpCell.getColumnIndex());
            cloneOldCell2NewCell(newCell, tmpCell);
        }
    }

    public static void cloneOldCell2NewCell(Cell newCell, Cell oldCell) {
        if (oldCell == null || newCell == null){return;}
        newCell.setCellStyle(oldCell.getCellStyle());
        switch (oldCell.getCellType()){
            case BOOLEAN -> newCell.setCellValue(oldCell.getBooleanCellValue());
            case NUMERIC -> newCell.setCellValue(oldCell.getNumericCellValue());
            case STRING -> newCell.setCellValue(oldCell.getStringCellValue());
            case FORMULA -> newCell.setCellValue(oldCell.getCellFormula());
            case ERROR -> newCell.setCellValue(oldCell.getErrorCellValue());
            case BLANK -> newCell.setCellValue("");
        }
    }

}

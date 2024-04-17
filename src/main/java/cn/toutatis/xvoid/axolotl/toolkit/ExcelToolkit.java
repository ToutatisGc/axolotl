package cn.toutatis.xvoid.axolotl.toolkit;

import cn.toutatis.xvoid.axolotl.Meta;
import cn.toutatis.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.components.SheetHeader;
import cn.toutatis.xvoid.axolotl.exceptions.AxolotlException;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkitKt;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.slf4j.Logger;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

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

    /**
     * 判断当前单元格是否在合并单元格中
     * @param sheet 工作表
     * @param rowIndex 行号
     * @param colIndex 列号
     * @return 当前单元格是否是合并单元格
     */
    public static CellRangeAddress isCellMerged(Sheet sheet, int rowIndex, int colIndex) {
        int numMergedRegions = sheet.getNumMergedRegions();
        for(int i = 0; i < numMergedRegions; i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if(rowIndex >= mergedRegion.getFirstRow() && rowIndex <= mergedRegion.getLastRow() &&
                    colIndex >= mergedRegion.getFirstColumn() && colIndex <= mergedRegion.getLastColumn()) {
                return mergedRegion;
            }
        }
        return null;
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
            case BOOLEAN:
                newCell.setCellValue(oldCell.getBooleanCellValue());
                break;
            case NUMERIC:
                newCell.setCellValue(oldCell.getNumericCellValue());
                break;
            case STRING:
                newCell.setCellValue(oldCell.getStringCellValue());
                break;
            case FORMULA:
                newCell.setCellValue(oldCell.getCellFormula());
                break;
            case ERROR:
                newCell.setCellValue(oldCell.getErrorCellValue());
                break;
            case BLANK:
                newCell.setBlank();
                break;
        }
    }

    /**
     * 创建行
     * @param sheet 工作表
     * @param row 行位置
     * @return 新行
     */
    public static Row createOrCatchRow(Sheet sheet, int row){
        Row writableRow = sheet.getRow(row);
        if (writableRow == null){
            writableRow = sheet.createRow(row);
        }
        return writableRow;
    }

    /**
     * 创建单元格
     * @param sheet 工作表
     * @param row 行位置
     * @param column 列位置
     * @param cellStyle 单元格样式
     * @return 新单元格
     */
    public static Cell createOrCatchCell(Sheet sheet, int row, int column, CellStyle cellStyle){
        Row writableRow = createOrCatchRow(sheet, row);
        Cell writableCell = writableRow.getCell(column);
        if (writableCell == null){
            writableCell = writableRow.createCell(column);
        }
        if (cellStyle != null){
            writableCell.setCellStyle(cellStyle);
        }
        return writableCell;
    }

    /**
     * 单元格赋值
     * @param sheet 工作表
     * @param row 行位置
     * @param column 列位置
     * @param cellStyle 单元格样式
     * @param value 值
     */
    public static void cellAssignment(Sheet sheet, int row, int column, CellStyle cellStyle ,Object value){
        Cell writableCell = createOrCatchCell(sheet, row, column, cellStyle);
        if (value != null){
            Class<?> valueClass = value.getClass();
            if ((ReflectToolkit.isWrapperClass(valueClass) && valueClass != String.class) || valueClass.isPrimitive()){
                writableCell.setCellValue((Double) value);
            }
            if (value instanceof String){
                writableCell.setCellValue((String) value);
            }else if (value instanceof Boolean){
                writableCell.setCellValue((Boolean) value);
            }else if (value instanceof Date){
                writableCell.setCellValue((Date) value);
            }else if (value instanceof LocalDateTime){
                writableCell.setCellValue((LocalDateTime) value);
            }else if (value instanceof LocalDate){
                writableCell.setCellValue((LocalDate) value);
            }else if (value instanceof Calendar){
                writableCell.setCellValue((Calendar) value);
            }else {
                throw new AxolotlException("不支持的写入类型");
            }
        }else {
            writableCell.setBlank();
        }
    }

    /**
     * 获取表头最大深度
     * @param headers 表头
     * @param depth 深度
     * @return 最大深度
     */
    public static int getMaxDepth(List<Header> headers, int depth) {
        int maxDepth = depth;
        for (Header header : headers) {
            if (header.getChilds() != null) {
                int subDepth = getMaxDepth(header.getChilds(), depth + 1);
                if (subDepth > maxDepth) {
                    maxDepth = subDepth;
                }
            }
        }
        return maxDepth;
    }

    /**
     * 获取实体类表头
     * @param clazz 类
     * @return 表头
     */
    public static List<Header> getHeaderList(Class<?> clazz) {
        ArrayList<Header> headerList = new ArrayList<>();
        List<Field> list = ReflectToolkit.getAllFields(clazz, true);
        for (Field field : list) {
            SheetHeader sheetHeader = field.getDeclaredAnnotation(SheetHeader.class);
            if (sheetHeader != null) {
                Header header = new Header(sheetHeader.name());
                header.setFieldName(field.getName());
                header.setColumnWidth(sheetHeader.width());
                headerList.add(header);
            }
        }
        return headerList;
    }

    /**
     * 创建下拉列表选项(单元格下拉框数据小于255字节时使用)
     *
     * @param sheet    所在Sheet页面
     * @param values   下拉框的选项值
     * @param firstRow 起始行（从0开始）
     * @param lastRow  终止行（从0开始）
     * @param firstCol 起始列（从0开始）
     * @param lastCol  终止列（从0开始）
     */
    public static void createDropDownList(Sheet sheet, String[] values, int firstRow, int lastRow, int firstCol, int lastCol) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidationConstraint constraint = helper.createExplicitListConstraint(values);
        DataValidation dataValidation = helper.createValidation(constraint, addressList);
        dataValidation.setSuppressDropDownArrow(true);
        dataValidation.setShowErrorBox(true);
        sheet.addValidationData(dataValidation);
    }

}

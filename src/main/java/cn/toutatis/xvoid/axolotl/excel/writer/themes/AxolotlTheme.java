package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;

public class AxolotlTheme extends AbstractInnerStyleRender implements ExcelStyleRender{

    private SXSSFWorkbook workbook;

    private static final IndexedColors THEME_COLOR = AxolotlCommendatoryColors.INDIGO;
    @Override
    public void renderHeader(SXSSFSheet sheet) {
        workbook = sheet.getWorkbook();
        this.createTitleRow(sheet);
        CellStyle headerCellStyle = this.createCommonCellStyle(THEME_COLOR, (short) 12,true);
        SXSSFRow columnNamesRow = sheet.createRow(1);
        columnNamesRow.setHeight((short) 400);
        List<String> columnNames = writerConfig.getColumnNames();
        for (int i = 0; i < columnNames.size(); i++) {
            SXSSFCell cell = columnNamesRow.createCell(i);
            cell.setCellValue(columnNames.get(i));
            cell.setCellStyle(headerCellStyle);
        }
    }

    private void createTitleRow(SXSSFSheet sheet){
//        sheet.setDefaultColumnStyle();
        SXSSFRow titleRow = sheet.createRow(0);
        titleRow.setHeight((short) 600);
        SXSSFCell startPositionCell = titleRow.createCell(0);
        startPositionCell.setCellValue(writerConfig.getTitle());
        CellStyle titleCellStyle = this.createCommonCellStyle(THEME_COLOR,(short) 18,true);
        startPositionCell.setCellStyle(titleCellStyle);
        CellRangeAddress cellAddresses = new CellRangeAddress(0, 0, 0, writerConfig.getColumnNames().size() - 1);
        for (int rowNum = cellAddresses.getFirstRow(); rowNum <= cellAddresses.getLastRow(); rowNum++) {
            Row currentRow = sheet.getRow(rowNum);
            if (currentRow == null) {currentRow = sheet.createRow(rowNum);}
            for (int colNum = cellAddresses.getFirstColumn(); colNum <= cellAddresses.getLastColumn(); colNum++) {
                Cell currentCell = currentRow.getCell(colNum);
                if (currentCell == null) {currentCell = currentRow.createCell(colNum);}
                currentCell.setCellStyle(titleCellStyle);
            }
        }
        sheet.addMergedRegion(cellAddresses);
    }

    private CellStyle createCommonCellStyle(IndexedColors color, short fontSize, boolean bold) {
        CellStyle commonCellStyle = workbook.createCellStyle();
        BorderStyle borderStyle = BorderStyle.MEDIUM;
        commonCellStyle.setBorderTop(borderStyle);
        commonCellStyle.setBorderRight(borderStyle);
        commonCellStyle.setBorderBottom(borderStyle);
        commonCellStyle.setBorderLeft(borderStyle);
        commonCellStyle.setTopBorderColor(color.getIndex());
        commonCellStyle.setRightBorderColor(color.getIndex());
        commonCellStyle.setBottomBorderColor(color.getIndex());
        commonCellStyle.setLeftBorderColor(color.getIndex());
        commonCellStyle.setAlignment(HorizontalAlignment.CENTER);
        commonCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setFontName("宋体");
        font.setBold(bold);
        font.setFontHeightInPoints(fontSize);
        commonCellStyle.setFont(font);
        return commonCellStyle;
    }
}

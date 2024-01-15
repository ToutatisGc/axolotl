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
        CellStyle headerCellStyle = StyleHelper.createCommonCellStyle(workbook, BorderStyle.MEDIUM,THEME_COLOR,true,"宋体",StyleHelper.STANDARD_TEXT_FONT_SIZE);
        SXSSFRow columnNamesRow = sheet.createRow(1);
        columnNamesRow.setHeight((short) 400);
        List<String> columnNames = writerConfig.getColumnNames();
        for (int i = 0; i < columnNames.size(); i++) {
            SXSSFCell cell = columnNamesRow.createCell(i);
            String name = columnNames.get(i);
            cell.setCellValue(name);
            cell.setCellStyle(headerCellStyle);
            sheet.setColumnWidth(i,StyleHelper.getPresetCellLength(name));
        }
        workbook.setPrintArea(workbook.getSheetIndex(sheet), 0,writerConfig.getColumnNames().size() - 1,0, 10 );
    }

    private void createTitleRow(SXSSFSheet sheet){
        SXSSFRow titleRow = sheet.createRow(0);
        titleRow.setHeight((short) 600);
        SXSSFCell startPositionCell = titleRow.createCell(0);
        startPositionCell.setCellValue(writerConfig.getTitle());
        CellStyle titleCellStyle = StyleHelper.createCommonCellStyle(workbook, BorderStyle.MEDIUM,THEME_COLOR,true,"宋体",StyleHelper.STANDARD_TITLE_FONT_SIZE);
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

}

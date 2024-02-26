package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractInnerStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AxolotlCommendatoryColors;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;
import java.util.Map;

public class AxolotlClassicalTheme extends AbstractInnerStyleRender implements ExcelStyleRender {

    private SXSSFWorkbook workbook;

    private static final IndexedColors THEME_COLOR = AxolotlCommendatoryColors.BLUE_GREY;

    private int alreadyWriteRow = -1;

    @Override
    public void renderHeader(SXSSFSheet sheet) {
        workbook = sheet.getWorkbook();
        this.createTitleRow(sheet);
        CellStyle headerCellStyle = StyleHelper.createCommonCellStyle(workbook, BorderStyle.MEDIUM,THEME_COLOR,true,"宋体",StyleHelper.STANDARD_TEXT_FONT_SIZE);
        SXSSFRow columnNamesRow = sheet.createRow(++alreadyWriteRow);
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

    @Override
    @SuppressWarnings("rawtypes")
    public void renderData(SXSSFSheet sheet,List<?> data) {
        CellStyle dataStyle = StyleHelper.createCommonCellStyle(
                sheet.getWorkbook(), BorderStyle.MEDIUM, THEME_COLOR, false,
                "宋体", StyleHelper.STANDARD_TEXT_FONT_SIZE
        );
        for (Object datum : data) {
            SXSSFRow dataRow = sheet.createRow(++alreadyWriteRow);
            dataRow.setHeight((short) 400);
            if (datum instanceof Map map){
                int colIdx = 0;
                for (Object o : map.keySet()) {
                    SXSSFCell cell = dataRow.createCell(colIdx);
                    Object dataObj = map.get(o);
                    String innerData = dataObj == null ? "" : dataObj.toString();
                    cell.setCellValue(innerData);
                    cell.setCellStyle(dataStyle);
                    colIdx++;
                }
            }
        }
    }

    private void createTitleRow(SXSSFSheet sheet){
        SXSSFRow titleRow = sheet.createRow(++alreadyWriteRow);
        titleRow.setHeight((short) 600);
        SXSSFCell startPositionCell = titleRow.createCell(0);
        startPositionCell.setCellValue(writerConfig.getTitle());
        CellStyle titleCellStyle = StyleHelper.createCommonCellStyle(workbook, BorderStyle.MEDIUM,THEME_COLOR,true,"宋体",StyleHelper.STANDARD_TITLE_FONT_SIZE);
        startPositionCell.setCellStyle(titleCellStyle);
        CellRangeAddress cellAddresses = new CellRangeAddress(0, 0, 0, writerConfig.getColumnNames().size() - 1);
        StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,titleCellStyle);
        sheet.addMergedRegion(cellAddresses);
    }

}

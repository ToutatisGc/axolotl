package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AxolotlCommendatoryColors;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AutoWriteContext;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.List;
import java.util.Map;

public class AxolotlClassicalTheme extends AbstractStyleRender implements ExcelStyleRender {

    private SXSSFWorkbook workbook;

    private static final IndexedColors THEME_COLOR = AxolotlCommendatoryColors.BLUE_GREY;

    private int alreadyWriteRow = -1;

    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        return null;
    }

    @Override
    public AxolotlWriteResult renderHeader(SXSSFSheet sheet) {
        // 1.渲染标题
        String title = writeConfig.getTitle();
        if(Validator.strNotBlank(title)){
            this.createTitleRow(sheet);
        }
        SXSSFRow columnNamesRow = sheet.createRow(++alreadyWriteRow);
        columnNamesRow.setHeight((short) 400);
        CellStyle headerCellStyle = StyleHelper.createStandardCellStyle(workbook, BorderStyle.MEDIUM,THEME_COLOR,true,"宋体",StyleHelper.STANDARD_TEXT_FONT_SIZE);
        List<Header> headers = context.getHeaders();


//        List<String> columnNames = writeConfig.getColumnNames();
//        for (int i = 0; i < columnNames.size(); i++) {
//            SXSSFCell cell = columnNamesRow.createCell(i);
//            String name = columnNames.get(i);
//            cell.setCellValue(name);
//            cell.setCellStyle(headerCellStyle);
//            sheet.setColumnWidth(i,StyleHelper.getPresetCellLength(name));
//        }
        return null;
    }

    private void renderHeaders(SXSSFSheet sheet,List<Header> headers){
        calculateHeaderLevel(headers,0);
    }

    private int calculateHeaderLevel(List<Header> headers,int level){
        if(headers != null && !headers.isEmpty()){
            for (Header header : headers) {
                List<Header> childs = header.getChilds();
                if(!childs.isEmpty()){
                    level++;
                    return calculateHeaderLevel(childs,level);
                }else{
                    return 0;
                }
            }
        }else{
            return level;
        }
        return level;
    }

    @Override
    @SuppressWarnings("rawtypes")
    public AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data) {
        CellStyle dataStyle = StyleHelper.createStandardCellStyle(
                sheet.getWorkbook(), BorderStyle.MEDIUM, THEME_COLOR, false,
                "宋体", StyleHelper.STANDARD_TEXT_FONT_SIZE
        );
//        List<Field> allFields = ReflectToolkit.getAllFields(datum.getClass(), true);
        for (Object datum : data) {
            SXSSFRow dataRow = sheet.createRow(++alreadyWriteRow);
            System.err.println("写入："+alreadyWriteRow+"="+datum);
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
        return null;
    }

    @Override
    public AxolotlWriteResult finish() {
        return null;
    }

    private void createTitleRow(SXSSFSheet sheet){
        SXSSFRow titleRow = sheet.createRow(++alreadyWriteRow);
        titleRow.setHeight((short) 600);
        SXSSFCell startPositionCell = titleRow.createCell(0);
        startPositionCell.setCellValue(writeConfig.getTitle());
        CellStyle titleCellStyle = StyleHelper.createStandardCellStyle(workbook, BorderStyle.MEDIUM,THEME_COLOR,true,"宋体",StyleHelper.STANDARD_TITLE_FONT_SIZE);
        startPositionCell.setCellStyle(titleCellStyle);
//        CellRangeAddress cellAddresses = new CellRangeAddress(0, 0, 0, writeConfig.getColumnNames().size() - 1);
//        StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,titleCellStyle);
//        sheet.addMergedRegion(cellAddresses);
    }

}

package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AxolotlCommendatoryColors;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AutoWriteContext;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;

import java.util.List;
import java.util.Map;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

public class AxolotlClassicalTheme extends AbstractStyleRender implements ExcelStyleRender {

    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlClassicalTheme.class);


    private static final IndexedColors THEME_COLOR = AxolotlCommendatoryColors.BLUE_GREY;

    private int alreadyWriteRow = -1;

    /**
     * 是否已经写入标题
     */
    private boolean alraedyWirteTitle = false;

    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        AxolotlWriteResult axolotlWriteResult;
        if(isFirstBatch()){
            axolotlWriteResult = new AxolotlWriteResult(true,"初始化成功");
            String sheetName = writeConfig.getSheetName();
            if(Validator.strNotBlank(sheetName)){
                int sheetIndex = writeConfig.getSheetIndex();
                info(LOGGER,"设置工作表索引[%s]表名为:[%s]",sheetIndex,sheetName);
                context.getWorkbook().setSheetName(sheetIndex,sheetName);
            }else {
                debug(LOGGER,"未设置工作表名称");
            }
        }else{
            axolotlWriteResult = new AxolotlWriteResult(true,"已初始化");
        }
        return axolotlWriteResult;
    }

    @Override
    public AxolotlWriteResult renderHeader(SXSSFSheet sheet) {
        // 1.渲染标题
        CellStyle titleRow = this.createTitleRow(sheet);
        // 2.渲染表头
        List<Header> headers = context.getHeaders();
        int headerMaxDepth = -1;
        int headerColumnCount = 0;
        if (headers != null && !headers.isEmpty()){
            alreadyWriteRow++;
            headerMaxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            //根节点渲染
            for (Header header : headers) {
                List<Header> childs = header.getChilds();
                int orlopCellNumber = header.countOrlopCellNumber();
                Row row = ExcelToolkit.createOrCatchRow(sheet, alreadyWriteRow);
                Cell cell = row.createCell(headerColumnCount, CellType.STRING);
                cell.setCellValue(header.getTitle());
                if (childs != null && !childs.isEmpty()){
                    CellRangeAddress cellAddresses = new CellRangeAddress(alreadyWriteRow, alreadyWriteRow, headerColumnCount, headerColumnCount+orlopCellNumber-1);
                    StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,titleRow);
                    sheet.addMergedRegion(cellAddresses);
                    recursionRenderHeaders(sheet,childs,headerMaxDepth,++alreadyWriteRow,headerColumnCount-1);
                }else{
                    CellRangeAddress cellAddresses = new CellRangeAddress(alreadyWriteRow, alreadyWriteRow+headerMaxDepth-1, headerColumnCount, headerColumnCount+orlopCellNumber-1);
                    StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,titleRow);
                    sheet.addMergedRegion(cellAddresses);
                }
                headerColumnCount+=orlopCellNumber;
            }
            System.err.println("maxDepth:"+headerMaxDepth);
        }else{
            debug(LOGGER,"未设置表头");
        }
        System.err.println(headerColumnCount);

        CellRangeAddress cellAddresses = new CellRangeAddress(0, 0, 0, headerColumnCount-1);
        StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,titleRow);
        sheet.addMergedRegion(cellAddresses);
//        if (headerMaxDepth > 0){
//
//        }
//        SXSSFRow columnNamesRow = sheet.createRow(++alreadyWriteRow);
//        columnNamesRow.setHeight((short) 400);
//        CellStyle headerCellStyle = StyleHelper.createStandardCellStyle(workbook, BorderStyle.MEDIUM,THEME_COLOR,true,"宋体",StyleHelper.STANDARD_TEXT_FONT_SIZE);
//        List<Header> headers = context.getHeaders();


//        List<String> columnNames = writeConfig.getColumnNames();
//        for (int i = 0; i < columnNames.size(); i++) {
//            SXSSFCell cell = columnNamesRow.createCell(i);
//            String name = columnNames.get(i);
//            cell.setCellValue(name);
//            cell.setCellStyle(headerCellStyle);
//            sheet.setColumnWidth(i,StyleHelper.getPresetCellLength(name));
//        }
        // 合并标题列

        return null;
    }

    private void recursionRenderHeaders(SXSSFSheet sheet,List<Header> headers,int maxLevel,int row,int column){
        if (headers != null && !headers.isEmpty()){
            for (Header header : headers) {
                System.err.println(header);
                int orlopCellNumber = header.countOrlopCellNumber();
                if (orlopCellNumber == 1){
                    Row row1 = ExcelToolkit.createOrCatchRow(sheet, row);
                    Cell cell = row1.createCell(column, CellType.STRING);
                    cell.setCellValue(header.getTitle());
                    sheet.setColumnWidth(column,StyleHelper.getPresetCellLength(header.getTitle()));
                    column++;
                    continue;
                }else{
                    if (header.getChilds() != null && !header.getChilds().isEmpty()){
                        recursionRenderHeaders(sheet,header.getChilds(),maxLevel,row++,column);
                    }else{

                    }
                }
            }
        }
//        calculateHeaderLevel(headers,0);
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

    private CellStyle createTitleRow(SXSSFSheet sheet){
        String title = writeConfig.getTitle();
        CellStyle cellStyle = null;
        if (Validator.strNotBlank(title)){
            debug(LOGGER,"设置工作表标题:[%s]",title);
            SXSSFRow titleRow = sheet.createRow(++alreadyWriteRow);
            titleRow.setHeight((short) 600);
            SXSSFCell startPositionCell = titleRow.createCell(0);
            startPositionCell.setCellValue(writeConfig.getTitle());
            cellStyle = StyleHelper.createStandardCellStyle(context.getWorkbook(), BorderStyle.MEDIUM,THEME_COLOR,true,"宋体",StyleHelper.STANDARD_TITLE_FONT_SIZE);
        }else{
            debug(LOGGER,"未设置工作表标题");
        }
        return cellStyle;
    }

}

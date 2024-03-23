package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Data;
import lombok.SneakyThrows;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.slf4j.Logger;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

public class AxolotlClassicalTheme extends AbstractStyleRender implements ExcelStyleRender {

    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlClassicalTheme.class);

    private static final AxolotlColor THEME_COLOR_XSSF = new AxolotlColor(68,114,199);

    private static final String FONT_NAME = "宋体";

    private Font MAIN_TEXT_FONT;

    private int alreadyWriteRow = -1;

    /**
     * 是否已经写入标题
     */
    private boolean alreadyWriteTitle = false;

    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        AxolotlWriteResult axolotlWriteResult;
        if(isFirstBatch()){
            MAIN_TEXT_FONT = StyleHelper.createWorkBookFont(context.getWorkbook(), FONT_NAME, false, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.BLACK);
            axolotlWriteResult = new AxolotlWriteResult(true,"初始化成功");
            String sheetName = writeConfig.getSheetName();
            if(Validator.strNotBlank(sheetName)){
                int sheetIndex = writeConfig.getSheetIndex();
                info(LOGGER,"设置工作表索引[%s]表名为:[%s]",sheetIndex,sheetName);
                context.getWorkbook().setSheetName(sheetIndex,sheetName);
            }else {
                debug(LOGGER,"未设置工作表名称");
            }

            CellStyle defaultStyle = StyleHelper.createStandardCellStyle(
                    context.getWorkbook(), BorderStyle.NONE, IndexedColors.WHITE, new AxolotlColor(255, 255, 255), MAIN_TEXT_FONT
            );
            // 将默认样式应用到所有单元格
            for (int i = 0; i < 26; i++) {
                sheet.setDefaultColumnStyle(i, defaultStyle);
                sheet.setDefaultColumnWidth(12);
            }
            sheet.setDefaultRowHeight((short) 350);
        }else{
            axolotlWriteResult = new AxolotlWriteResult(true,"已初始化");
        }
        return axolotlWriteResult;
    }

    /**
     * 表头递归信息
     */
    @Data
    public class HeaderRecursiveInfo implements Serializable,Cloneable{

        /**
         * 渲染的总行数
         */
        private int allRow;

        /**
         * 起始列
         */
        private int startColumn;

        /**
         * 已经写入的列
         */
        private int alreadyWriteColumn;

        /**
         * 渲染的单元格样式
         */
        private CellStyle cellStyle;

        /**
         * 渲染的行高度
         */
        private short rowHeight;
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
            Font font = StyleHelper.createWorkBookFont(context.getWorkbook(), FONT_NAME, true, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.WHITE);
            CellStyle headerCellStyle = StyleHelper.createStandardCellStyle(
                    context.getWorkbook(), BorderStyle.MEDIUM, IndexedColors.WHITE, THEME_COLOR_XSSF,font
            );
            alreadyWriteRow++;
            headerMaxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            debug(LOGGER,"起始行次为[%s]，表头最大深度为[%s]",alreadyWriteRow,headerMaxDepth);
            //根节点渲染
            for (Header header : headers) {
                int startRow = alreadyWriteRow;
                Row row = ExcelToolkit.createOrCatchRow(sheet, startRow);
                row.setHeight(StyleHelper.STANDARD_HEADER_ROW_HEIGHT);
                Cell cell = row.createCell(headerColumnCount, CellType.STRING);
                String title = header.getName();
                cell.setCellValue(title);
                int orlopCellNumber = header.countOrlopCellNumber();
                context.setAlreadyWrittenColumns(context.getAlreadyWrittenColumns()+orlopCellNumber);
                debug(LOGGER,"渲染表头[%s],行[%s],列[%s],子表头列数量[%s]",title,startRow,headerColumnCount,orlopCellNumber);
                // 有子节点说明需要向下迭代并合并
                CellRangeAddress cellAddresses;
                if (header.getChilds()!=null && !header.getChilds().isEmpty()){
                    List<Header> childs = header.getChilds();
                    int childMaxDepth = ExcelToolkit.getMaxDepth(childs, 0);
                    cellAddresses = new CellRangeAddress(startRow, startRow+(headerMaxDepth-childMaxDepth)-1, headerColumnCount, headerColumnCount+orlopCellNumber-1);
                    HeaderRecursiveInfo headerRecursiveInfo = new HeaderRecursiveInfo();
                    headerRecursiveInfo.setAllRow(alreadyWriteRow+headerMaxDepth+1);
                    headerRecursiveInfo.setStartColumn(headerColumnCount);
                    headerRecursiveInfo.setAlreadyWriteColumn(headerColumnCount);
                    headerRecursiveInfo.setCellStyle(headerCellStyle);
                    headerRecursiveInfo.setRowHeight(StyleHelper.STANDARD_HEADER_ROW_HEIGHT);
                    recursionRenderHeaders(sheet,childs, headerRecursiveInfo);
                }else{
                    cellAddresses = new CellRangeAddress(startRow, (startRow+headerMaxDepth)-1, headerColumnCount, headerColumnCount);
                }
                StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,headerCellStyle);
                if (headerMaxDepth > 1){
                    sheet.addMergedRegion(cellAddresses);
                }
                headerColumnCount+=orlopCellNumber;
            }
        }else{
            debug(LOGGER,"未设置表头");
        }
        alreadyWriteRow+=(headerMaxDepth-1);
        debug(LOGGER,"合并标题栏单元格,共[%s]列",headerColumnCount);
        CellRangeAddress cellAddresses = new CellRangeAddress(0, 0, 0, headerColumnCount-1);
        StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,titleRow);
        if (headerColumnCount > 1){
            sheet.addMergedRegion(cellAddresses);
        }
        sheet.createFreezePane(0, alreadyWriteRow+1);

        return null;
    }

    @SneakyThrows
    private void recursionRenderHeaders(SXSSFSheet sheet, List<Header> headers, HeaderRecursiveInfo headerRecursiveInfo){
        if (headers != null && !headers.isEmpty()){
            int maxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            int startRow = headerRecursiveInfo.getAllRow() - maxDepth -1;
            Row row = ExcelToolkit.createOrCatchRow(sheet,startRow);
            row.setHeight(headerRecursiveInfo.getRowHeight());
            for (Header header : headers) {
                int alreadyWriteColumn = headerRecursiveInfo.getAlreadyWriteColumn();
                Cell cell = ExcelToolkit.createOrCatchCell(sheet, row.getRowNum(), alreadyWriteColumn, null);
                cell.setCellValue(header.getName());
                int childCount = header.countOrlopCellNumber();
                int endColumnPosition = (alreadyWriteColumn + childCount);
                if (header.getChilds()!=null && !header.getChilds().isEmpty()){
                    CellRangeAddress cellAddresses = new CellRangeAddress(startRow, startRow, alreadyWriteColumn, endColumnPosition-1);
                    StyleHelper.renderMergeRegionStyle(sheet,cellAddresses, headerRecursiveInfo.getCellStyle());
                    sheet.addMergedRegion(cellAddresses);
                }else{
                    cell.setCellStyle(headerRecursiveInfo.getCellStyle());
                }
                headerRecursiveInfo.setAlreadyWriteColumn(endColumnPosition);
                headerRecursiveInfo.setStartColumn(alreadyWriteColumn);
                if (header.getChilds() != null && !header.getChilds().isEmpty()){
                    HeaderRecursiveInfo child = new HeaderRecursiveInfo();
                    BeanUtils.copyProperties(child, headerRecursiveInfo);
                    child.setAlreadyWriteColumn(headerRecursiveInfo.getStartColumn());
                    recursionRenderHeaders(sheet,header.getChilds(),child);
                }
            }
        }
    }

    @Override
    @SuppressWarnings("rawtypes")
    public AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data) {
        SXSSFWorkbook workbook = context.getWorkbook();
        BorderStyle borderStyle = BorderStyle.THIN;
        IndexedColors borderColor = IndexedColors.WHITE;
        CellStyle dataStyle = StyleHelper.createCellStyle(
                workbook, borderStyle, borderColor, new AxolotlColor(217,226,243),MAIN_TEXT_FONT
        );
        CellStyle dataStyleOdd = StyleHelper.createCellStyle(
                workbook ,borderStyle , borderColor, new AxolotlColor(181,197,230),MAIN_TEXT_FONT
        );
        DataFormat dataFormat = workbook.createDataFormat();
        short textFormatIndex = dataFormat.getFormat("@");
        dataStyle.setDataFormat(textFormatIndex);
        dataStyleOdd.setDataFormat(textFormatIndex);
        for (int i = 0, dataSize = data.size(); i < dataSize; i++) {
            boolean isOdd = i % 2 == 0;
            Object datum = data.get(i);
            SXSSFRow dataRow = sheet.createRow(++alreadyWriteRow);
            System.err.println("写入：" + alreadyWriteRow + "=" + datum);
            dataRow.setHeight((short) 400);
            if (datum instanceof Map map) {
                int colIdx = 0;
                for (Object o : map.keySet()) {
                    SXSSFCell cell = dataRow.createCell(colIdx);
                    Object dataObj = map.get(o);
                    String innerData = dataObj == null ? "" : dataObj.toString();
                    cell.setCellValue(innerData);
                    if (isOdd){
                        cell.setCellStyle(dataStyleOdd);
                    }else{
                        cell.setCellStyle(dataStyle);
                    }

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
            titleRow.setHeight(StyleHelper.STANDARD_TITLE_ROW_HEIGHT);
            SXSSFCell startPositionCell = titleRow.createCell(0);
            startPositionCell.setCellValue(writeConfig.getTitle());
            Font font = StyleHelper.createWorkBookFont(context.getWorkbook(), FONT_NAME, true, StyleHelper.STANDARD_TITLE_FONT_SIZE, IndexedColors.WHITE);
            cellStyle = StyleHelper.createStandardCellStyle(
                    context.getWorkbook(), BorderStyle.THICK, IndexedColors.WHITE, THEME_COLOR_XSSF,font
            );
        }else{
            debug(LOGGER,"未设置工作表标题");
        }
        return cellStyle;
    }

}

package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AxolotlCommendatoryColors;
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
import org.slf4j.Logger;

import java.io.Serializable;
import java.util.List;
import java.util.Map;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

public class AxolotlClassicalTheme extends AbstractStyleRender implements ExcelStyleRender {

    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlClassicalTheme.class);

    private static final IndexedColors THEME_COLOR = AxolotlCommendatoryColors.BLUE_GREY;

    public static final int TITLE_FONT_SIZE = 18;

    private int alreadyWriteRow = -1;

    /**
     * 是否已经写入标题
     */
    private boolean alreadyWriteTitle = false;

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
            alreadyWriteRow++;
            headerMaxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            debug(LOGGER,"起始行次为[%s]，表头最大深度为[%s]",alreadyWriteRow,headerMaxDepth);
            //根节点渲染
            for (Header header : headers) {
                int startRow = alreadyWriteRow;

                Row row = ExcelToolkit.createOrCatchRow(sheet, startRow);
                Cell cell = row.createCell(headerColumnCount, CellType.STRING);
                String title = header.getName();
                cell.setCellValue(title);
                int orlopCellNumber = header.countOrlopCellNumber();
                debug(LOGGER,"渲染表头[%s],行[%s],列[%s],子表头列数量[%s]",title,startRow,headerColumnCount,orlopCellNumber);
                // 有子节点说明需要向下迭代并合并
                CellRangeAddress cellAddresses;
                if (orlopCellNumber > 1){
                    List<Header> childs = header.getChilds();
                    int childMaxDepth = ExcelToolkit.getMaxDepth(childs, 0);
                    cellAddresses = new CellRangeAddress(startRow, startRow+(headerMaxDepth-childMaxDepth)-1, headerColumnCount, headerColumnCount+orlopCellNumber-1);
                    HeaderRecursiveInfo headerRecursiveInfo = new HeaderRecursiveInfo();
                    headerRecursiveInfo.setAllRow(alreadyWriteRow+headerMaxDepth+1);
                    headerRecursiveInfo.setStartColumn(headerColumnCount);
                    headerRecursiveInfo.setAlreadyWriteColumn(headerColumnCount);
                    headerRecursiveInfo.setCellStyle(titleRow);
                    recursionRenderHeaders(sheet,childs, headerRecursiveInfo);
                }else{
                    cellAddresses = new CellRangeAddress(startRow, (startRow+headerMaxDepth)-1, headerColumnCount, headerColumnCount);
                }
                StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,titleRow);
                if (headerMaxDepth > 1){
                    sheet.addMergedRegion(cellAddresses);
                }
                headerColumnCount+=orlopCellNumber;
            }
        }else{
            debug(LOGGER,"未设置表头");
        }
        debug(LOGGER,"合并标题栏单元格,共[%s]列",headerColumnCount);
        CellRangeAddress cellAddresses = new CellRangeAddress(0, 0, 0, headerColumnCount-1);
        StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,titleRow);
        if (headerColumnCount > 1){
            sheet.addMergedRegion(cellAddresses);
        }

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

    @SneakyThrows
    private void recursionRenderHeaders(SXSSFSheet sheet, List<Header> headers, HeaderRecursiveInfo headerRecursiveInfo){
        if (headers != null && !headers.isEmpty()){
            int maxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            int startRow = headerRecursiveInfo.getAllRow() - maxDepth -1;
            Row row = ExcelToolkit.createOrCatchRow(sheet,startRow);
            row.setHeight((short) 300);
            for (Header header : headers) {
                int alreadyWriteColumn = headerRecursiveInfo.getAlreadyWriteColumn();
                Cell cell = ExcelToolkit.createOrCatchCell(sheet, row.getRowNum(), alreadyWriteColumn, null);
                cell.setCellValue(header.getName());
                int childCount = header.countOrlopCellNumber();
                int endColumnPosition = (alreadyWriteColumn + childCount);
                if (childCount > 1){
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

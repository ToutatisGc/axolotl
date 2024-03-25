package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
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
import java.lang.reflect.Field;
import java.util.*;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

public class AxolotlClassicalTheme extends AbstractStyleRender implements ExcelStyleRender {

    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlClassicalTheme.class);

    private static final AxolotlColor THEME_COLOR_XSSF = new AxolotlColor(68,114,199);

    private static final String FONT_NAME = "宋体";

    private Font MAIN_TEXT_FONT;

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
        int alreadyWriteRow = context.getAlreadyWriteRow();
        if (headers != null && !headers.isEmpty()){
            List<Header> tmpHeaders;
            if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER)){
                tmpHeaders = new ArrayList<>();
                tmpHeaders.add(new Header("序号"));
                tmpHeaders.addAll(headers);
            }else{
                tmpHeaders = headers;
            }
            Font font = StyleHelper.createWorkBookFont(context.getWorkbook(), FONT_NAME, true, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.WHITE);
            CellStyle headerCellStyle = StyleHelper.createStandardCellStyle(
                    context.getWorkbook(), BorderStyle.MEDIUM, IndexedColors.WHITE, THEME_COLOR_XSSF,font
            );
            context.setAlreadyWriteRow(++alreadyWriteRow);
            headerMaxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            debug(LOGGER,"起始行次为[%s]，表头最大深度为[%s]",alreadyWriteRow,headerMaxDepth);
            //根节点渲染
            for (Header header : tmpHeaders) {
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
        context.setAlreadyWriteRow(alreadyWriteRow);
        context.setAlreadyWrittenColumns(headerColumnCount);
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
                CellRangeAddress cellAddresses;
                int mergeRowNumber = startRow + maxDepth - 1;
                if (header.getChilds()!=null && !header.getChilds().isEmpty()){
                    cellAddresses = new CellRangeAddress(startRow, startRow, alreadyWriteColumn, endColumnPosition-1);
                }else{
                    cellAddresses = new CellRangeAddress(startRow, startRow + maxDepth-1, alreadyWriteColumn, endColumnPosition-1);
                    cell.setCellStyle(headerRecursiveInfo.getCellStyle());
                }
                StyleHelper.renderMergeRegionStyle(sheet,cellAddresses, headerRecursiveInfo.getCellStyle());
                if (mergeRowNumber !=  startRow){
                    sheet.addMergedRegion(cellAddresses);
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
    @SuppressWarnings({"rawtypes","unchecked"})
    public AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data) {
        SXSSFWorkbook workbook = context.getWorkbook();
        BorderStyle borderStyle = BorderStyle.THIN;
        IndexedColors borderColor = IndexedColors.WHITE;
        // 交叉样式
        DataFormat dataFormat = workbook.createDataFormat();
        short textFormatIndex = dataFormat.getFormat("@");
        CellStyle dataStyle = StyleHelper.createStandardCellStyle(workbook, borderStyle, borderColor, new AxolotlColor(217,226,243),MAIN_TEXT_FONT);
        CellStyle dataStyleOdd = StyleHelper.createStandardCellStyle(workbook ,borderStyle , borderColor, new AxolotlColor(181,197,230),MAIN_TEXT_FONT);
        dataStyle.setDataFormat(textFormatIndex);
        dataStyleOdd.setDataFormat(textFormatIndex);
        boolean autoInsertSerialNumber = writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER);
        for (int i = 0, dataSize = data.size(); i < dataSize; i++) {
            // 获取对象属性
            Object dataObj = data.get(i);
            HashMap<String, Object> dataMap = new LinkedHashMap<>();
            if (dataObj instanceof Map map) {
                dataMap.putAll(map);
            }else{
                List<Field> fields = ReflectToolkit.getAllFields(dataObj.getClass(), true);
                fields.forEach(field -> {
                    field.setAccessible(true);
                    String fieldName = field.getName();
                    try {
                        dataMap.put(fieldName,field.get(dataObj));
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                        throw new AxolotlWriteException("获取对象字段错误");
                    }
                });
            }
            // 初始化内容
            CellStyle innerStyle = i % 2 == 0 ? dataStyle : dataStyleOdd;
            HashMap<Integer, Integer> writtenColumnMap = new HashMap<>();
            int alreadyWriteRow = context.getAlreadyWriteRow();
            context.setAlreadyWriteRow(++alreadyWriteRow);
            SXSSFRow dataRow = sheet.createRow(alreadyWriteRow);
            int writtenColumn = 0;
            int serialNumber = context.getAndIncrementSerialNumber();
            // 写入序号
            if (autoInsertSerialNumber){
                SXSSFCell cell = dataRow.createCell(writtenColumn);
                cell.setCellValue(serialNumber);
                cell.setCellStyle(innerStyle);
                writtenColumnMap.put(writtenColumn++,1);
            }
            // 写入数据
            for (Map.Entry<String, Object> dataEntry : dataMap.entrySet()) {
                Object value = dataEntry.getValue();
                SXSSFCell cell = dataRow.createCell(writtenColumn);
                if (value == null){
                    cell.setCellValue(writeConfig.getBlankValue());
                }else{
                    cell.setCellValue(value.toString());
                }
                cell.setCellStyle(innerStyle);
                writtenColumnMap.put(writtenColumn++,1);
            }
            for (int alreadyColumnIdx = 0; alreadyColumnIdx < context.getAlreadyWrittenColumns(); alreadyColumnIdx++) {
                if (!writtenColumnMap.containsKey(alreadyColumnIdx)){
                    SXSSFCell cell = dataRow.createCell(alreadyColumnIdx);
                    cell.setBlank();
                    cell.setCellStyle(innerStyle);
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
            int alreadyWriteRow = context.getAlreadyWriteRow();
            context.setAlreadyWriteRow(++alreadyWriteRow);
            SXSSFRow titleRow = sheet.createRow(alreadyWriteRow);
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

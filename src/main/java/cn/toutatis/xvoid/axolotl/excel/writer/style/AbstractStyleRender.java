package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AutoWriteContext;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlConstant;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import lombok.SneakyThrows;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.*;

import static cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper.START_POSITION;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

/**
 * 样式渲染器抽象类
 * 继承此抽象类可以获取环境变量实现自定义样式渲染
 * @author Toutatis_Gc
 */
@Getter
public abstract class AbstractStyleRender implements ExcelStyleRender{

    @Setter
    protected AutoWriteConfig writeConfig;

    @Setter
    protected AutoWriteContext context;

    private Logger LOGGER;

    /**
     * 字体名称
     */
    private String fontName;

    /**
     * 数据写入已进行错误提示
     */
    private boolean alreadyNotice = false;

    public AbstractStyleRender(Logger LOGGER) {
        this.LOGGER = LOGGER;
        this.fontName = StyleHelper.STANDARD_FONT_NAME;
    }

    public AbstractStyleRender(Logger LOGGER,String fontName) {
        this.LOGGER = LOGGER;
        this.fontName = fontName;
    }

    public static final String TOTAL_HEADER_COUNT_KEY = "";

    /**
     * 是否是第一批次数据
     * @return true/false
     */
    public boolean isFirstBatch(){
        return context.isFirstBatch(context.getSwitchSheetIndex());
    }

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
            boolean fillWhite = writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_FILL_DEFAULT_CELL_WHITE);
            if (fillWhite){
                fillWhiteCell(sheet,fontName);
            }
        }else {
            axolotlWriteResult = new AxolotlWriteResult(true,"已初始化");
        }
        return axolotlWriteResult;
    }

    /**
     * 填充空白单元格
     * @param sheet 工作表
     * @param fontName 字体名称
     */
    public void fillWhiteCell(Sheet sheet,String fontName){
        Font font = StyleHelper.createWorkBookFont(context.getWorkbook(), fontName, false, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.BLACK);
        CellStyle defaultStyle = StyleHelper.createStandardCellStyle(
                context.getWorkbook(), BorderStyle.NONE, IndexedColors.WHITE, new AxolotlColor(255, 255, 255), font
        );
        // 将默认样式应用到所有单元格
        for (int i = 0; i < 26; i++) {
            sheet.setDefaultColumnStyle(i, defaultStyle);
            sheet.setDefaultColumnWidth(12);
        }
        sheet.setDefaultRowHeight((short) 400);
    }

    /**
     * Part.1 表头
     * Step.1 创建标题行
     * @param sheet 工作表
     * @return 渲染结果
     */
    public AxolotlWriteResult createTitleRow(SXSSFSheet sheet){
        String title = writeConfig.getTitle();
        if (Validator.strNotBlank(title)){
            debug(LOGGER,"设置工作表标题:[%s]",title);
            int switchSheetIndex = context.getSwitchSheetIndex();
            Map<Integer, Integer> alreadyWriteRowMap = context.getAlreadyWriteRow();
            int alreadyWriteRow = alreadyWriteRowMap.getOrDefault(switchSheetIndex,-1);
            alreadyWriteRowMap.put(switchSheetIndex,++alreadyWriteRow);
            SXSSFRow titleRow = sheet.createRow(alreadyWriteRow);
            titleRow.setHeight(StyleHelper.STANDARD_TITLE_ROW_HEIGHT);
            SXSSFCell startPositionCell = titleRow.createCell(START_POSITION);
            startPositionCell.setCellValue(writeConfig.getTitle());
            return new AxolotlWriteResult(true, LoggerHelper.format("设置工作表标题:[%s]",title));
        }else{
            String message = "未设置工作表标题";
            debug(LOGGER,message);
            return new AxolotlWriteResult(false, message);
        }
    }

    /**
     *  Part.1 表头
     *  Step.3 合并标题栏单元格并赋予样式
     * @param sheet 工作表
     */
    public void mergeTitleRegion(SXSSFSheet sheet,int titleColumnCount,CellStyle titleStyle){
        if (titleColumnCount > 1){
            debug(LOGGER,"合并标题栏单元格,共[%s]列",titleColumnCount);
            CellRangeAddress cellAddresses = new CellRangeAddress(START_POSITION, START_POSITION, START_POSITION, titleColumnCount-1);
            StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,titleStyle);
            sheet.addMergedRegion(cellAddresses);
        }else{
            sheet.getRow(START_POSITION).getCell(START_POSITION).setCellStyle(titleStyle);
        }
    }

    /**
     * Part.1 表头
     * Step.2 递归表头
     * @param sheet 工作表
     * @param headerDefaultCellStyle 表头默认样式
     */
    public AxolotlWriteResult defaultRenderHeaders(SXSSFSheet sheet, CellStyle headerDefaultCellStyle){
        int switchSheetIndex = context.getSwitchSheetIndex();
        List<Header> headers = context.getHeaders().get(switchSheetIndex);
        int headerMaxDepth = -1;
        int headerColumnCount = 0;
        int alreadyWriteRow = context.getAlreadyWriteRow().getOrDefault(context.getSwitchSheetIndex(),-1);
        if (headers != null && !headers.isEmpty()){
            List<Header> cacheHeaders;
            if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER)){
                cacheHeaders = new ArrayList<>();
                cacheHeaders.add(new Header("序号"));
                cacheHeaders.addAll(headers);
            }else{
                cacheHeaders = headers;
            }
            context.getAlreadyWriteRow().put(switchSheetIndex,++alreadyWriteRow);
            headerMaxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            debug(LOGGER,"起始行次为[%s]，表头最大深度为[%s]",alreadyWriteRow,headerMaxDepth);
            int sheetIndex = writeConfig.getSheetIndex();
            Map<String, Integer> headerCache = context.getHeaderColumnIndexMapping().row(sheetIndex);
            //根节点渲染
            for (Header header : cacheHeaders) {
                CellStyle usedCellStyle = headerDefaultCellStyle;
                usedCellStyle = getCellStyle(header, usedCellStyle);
                Row row = ExcelToolkit.createOrCatchRow(sheet, alreadyWriteRow);
                row.setHeight(StyleHelper.STANDARD_HEADER_ROW_HEIGHT);

                Cell cell = row.createCell(headerColumnCount, CellType.STRING);
                String title = header.getName();
                cell.setCellValue(title);
                int orlopCellNumber = header.countOrlopCellNumber();
                context.getAlreadyWrittenColumns().put(switchSheetIndex,context.getAlreadyWrittenColumns().getOrDefault(switchSheetIndex,0)+orlopCellNumber);
                debug(LOGGER,"渲染表头[%s],行[%s],列[%s],子表头列数量[%s]",title, alreadyWriteRow,headerColumnCount,orlopCellNumber);
                // 有子节点说明需要向下迭代并合并
                CellRangeAddress cellAddresses;
                if (header.getChilds()!=null && !header.getChilds().isEmpty()){
                    List<Header> childs = header.getChilds();
                    int childMaxDepth = ExcelToolkit.getMaxDepth(childs, 0);
                    cellAddresses = new CellRangeAddress(alreadyWriteRow, alreadyWriteRow +(headerMaxDepth-childMaxDepth)-1, headerColumnCount, headerColumnCount+orlopCellNumber-1);
                    HeaderRecursiveInfo headerRecursiveInfo = new HeaderRecursiveInfo();
                    headerRecursiveInfo.setAllRow(alreadyWriteRow+headerMaxDepth+1);
                    headerRecursiveInfo.setStartColumn(headerColumnCount);
                    headerRecursiveInfo.setAlreadyWriteColumn(headerColumnCount);
                    headerRecursiveInfo.setCellStyle(headerDefaultCellStyle);
                    headerRecursiveInfo.setRowHeight(StyleHelper.STANDARD_HEADER_ROW_HEIGHT);
                    recursionRenderHeaders(sheet,childs, headerRecursiveInfo);
                }else{
                    cellAddresses = new CellRangeAddress(alreadyWriteRow, (alreadyWriteRow +headerMaxDepth)-1, headerColumnCount, headerColumnCount);
                    int columnWidth = header.getColumnWidth();
                    if (columnWidth == -1){
                        columnWidth = StyleHelper.getPresetCellLength(title);
                    }
                    String fieldName = header.getFieldName();
                    if (fieldName != null){
                        debug(LOGGER,"映射字段[%s]到列索引[%s]",fieldName,headerColumnCount);
                        headerCache.put(fieldName,headerColumnCount);
                    }
                    debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",headerColumnCount,header.getName(),columnWidth);
                    sheet.setColumnWidth(headerColumnCount, columnWidth);
                }
                StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,usedCellStyle);
                if (headerMaxDepth > 1){
                    sheet.addMergedRegion(cellAddresses);
                }
                headerColumnCount+=orlopCellNumber;
            }
        }else{
            headerMaxDepth = 0;
            debug(LOGGER,"未设置表头");
        }
        context.getHeaderRowCount().put(switchSheetIndex,headerMaxDepth);
        alreadyWriteRow+=(headerMaxDepth-1);
        context.getAlreadyWriteRow().put(switchSheetIndex,alreadyWriteRow);
        context.getAlreadyWrittenColumns().put(switchSheetIndex,headerColumnCount);
        return new AxolotlWriteResult(true, "渲染表头成功");
    }

    /**
     * Part.1 表头
     * 辅助方法 获取表头Header样式
     * @param header 表头
     * @param usedCellStyle 使用样式
     * @return 表头样式
     */
    public CellStyle getCellStyle(Header header, CellStyle usedCellStyle) {
        if (header.getCustomCellStyle() != null){
            usedCellStyle = header.getCustomCellStyle();
        }else{
            AxolotlCellStyle axolotlCellStyle = header.getAxolotlCellStyle();
            if (axolotlCellStyle != null){
                Font axolotlCustomFont = StyleHelper.createWorkBookFont(
                        context.getWorkbook(),
                        axolotlCellStyle.getFontName(),
                        axolotlCellStyle.isFontBold(),
                        axolotlCellStyle.getFontSize(),
                        axolotlCellStyle.getFontColor());
                usedCellStyle = StyleHelper.createStandardCellStyle(
                        context.getWorkbook(),
                        axolotlCellStyle.getBorderStyle(),
                        axolotlCellStyle.getBorderColor(),
                        axolotlCellStyle.getForegroundColor(),
                        axolotlCustomFont
                );
            }
        }
        return usedCellStyle;
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

    /**
     * 递归渲染表头
     * @param sheet 工作表
     * @param headers 表头集合
     * @param headerRecursiveInfo 递归信息
     */
    @SneakyThrows
    private void recursionRenderHeaders(SXSSFSheet sheet, List<Header> headers, HeaderRecursiveInfo headerRecursiveInfo){
        if (headers != null && !headers.isEmpty()){
            int maxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            int startRow = headerRecursiveInfo.getAllRow() - maxDepth -1;
            Row row = ExcelToolkit.createOrCatchRow(sheet,startRow);
            row.setHeight(headerRecursiveInfo.getRowHeight());
            int sheetIndex = writeConfig.getSheetIndex();
            Map<String, Integer> headerCache = context.getHeaderColumnIndexMapping().row(sheetIndex);
            for (Header header : headers) {
                CellStyle usedCellStyle = headerRecursiveInfo.getCellStyle();
                usedCellStyle = getCellStyle(header, usedCellStyle);
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
                    int columnWidth = header.getColumnWidth();
                    if (columnWidth == -1){
                        columnWidth = StyleHelper.getPresetCellLength(header.getName());
                    }
                    debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",alreadyWriteColumn,header.getName(),columnWidth);
                    sheet.setColumnWidth(alreadyWriteColumn, columnWidth);
                }
                StyleHelper.renderMergeRegionStyle(sheet,cellAddresses, usedCellStyle);
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
                }else{
                    String fieldName = header.getFieldName();
                    if (fieldName != null){
                        debug(LOGGER,"映射字段[%s]到列索引[%s]",fieldName,alreadyWriteColumn);
                        headerCache.put(fieldName,alreadyWriteColumn);
                    }
                }
            }
        }
    }


    /**
     * 默认行为渲染数据
     * @param sheet 工作表
     * @return 写入结果
     */
    @SuppressWarnings({"rawtypes","unchecked"})
    public void defaultRenderNextData(SXSSFSheet sheet,Object data,CellStyle rowStyle){
        // 获取对象属性
        HashMap<String, Object> dataMap = new LinkedHashMap<>();
        if (data instanceof Map map) {
            dataMap.putAll(map);
        }else{
            List<Field> fields = ReflectToolkit.getAllFields(data.getClass(), true);
            fields.forEach(field -> {
                field.setAccessible(true);
                String fieldName = field.getName();
                try {
                    dataMap.put(fieldName,field.get(data));
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                    throw new AxolotlWriteException("获取对象字段错误");
                }
            });
        }
        // 初始化内容
        HashMap<Integer, Integer> writtenColumnMap = new HashMap<>();
        int switchSheetIndex = getContext().getSwitchSheetIndex();
        Map<Integer, Integer> alreadyWriteRowMap = context.getAlreadyWriteRow();
        int alreadyWriteRow = alreadyWriteRowMap.getOrDefault(switchSheetIndex,-1);
        alreadyWriteRowMap.put(switchSheetIndex,++alreadyWriteRow);
        SXSSFRow dataRow = sheet.createRow(alreadyWriteRow);
        int writtenColumn = START_POSITION;
        int serialNumber = context.getAndIncrementSerialNumber() - context.getHeaderRowCount().get(switchSheetIndex);
        // 写入序号
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER)){
            SXSSFCell cell = dataRow.createCell(writtenColumn);
            cell.setCellValue(serialNumber);
            cell.setCellStyle(rowStyle);
            writtenColumnMap.put(writtenColumn++,1);
        }
        // 写入数据
        Map<String, Integer> columnMapping = context.getHeaderColumnIndexMapping().row(context.getSwitchSheetIndex());
        Map<Integer, Integer> unmappedColumnCount =  new HashMap<>();
        columnMapping.forEach((key, value) -> unmappedColumnCount.put(value, 1));
        boolean columnMappingEmpty = columnMapping.isEmpty();
        boolean useOrderField = true;
        for (Map.Entry<String, Object> dataEntry : dataMap.entrySet()) {
            SXSSFCell cell;
            if (columnMappingEmpty){
                cell = dataRow.createCell(writtenColumn);
            }else{
                useOrderField = false;
                if (columnMapping.containsKey(dataEntry.getKey())){
                    cell = (SXSSFCell) ExcelToolkit.createOrCatchCell(sheet,alreadyWriteRow,columnMapping.get(dataEntry.getKey()),null);
                }else {
                    if (!alreadyNotice){
                        warn(LOGGER,"未映射字段[%s]请在表头Header中映射字段!",dataEntry.getKey());
                        alreadyNotice = true;
                    }
                    continue;
                }
            }
            Object value = dataEntry.getValue();
            if (value == null){
                cell.setCellValue(writeConfig.getBlankValue());
            }else{
                String valueString = value.toString();
                if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING) && Validator.strIsNumber(valueString)){
                    Map<Integer, BigDecimal> endingTotalMapping = context.getEndingTotalMapping().row(switchSheetIndex);
                    int columnIndex = cell.getColumnIndex();
                    if (endingTotalMapping.containsKey(columnIndex)){
                        BigDecimal newValue = endingTotalMapping.get(columnIndex).add(BigDecimal.valueOf(Double.parseDouble(valueString)));
                        endingTotalMapping.put(columnIndex,newValue);
                    }else{
                        endingTotalMapping.put(columnIndex,BigDecimal.valueOf(Double.parseDouble(valueString)));
                    }
                }
                cell.setCellValue(valueString);
            }
            unmappedColumnCount.remove(cell.getColumnIndex());
            cell.setCellStyle(rowStyle);
            writtenColumnMap.put(writtenColumn++,1);
        }
        for (int alreadyColumnIdx = 0; alreadyColumnIdx < context.getAlreadyWrittenColumns().get(switchSheetIndex); alreadyColumnIdx++) {
            SXSSFCell cell = null;
            if (useOrderField){
                if (!writtenColumnMap.containsKey(alreadyColumnIdx)){
                    cell = dataRow.createCell(alreadyColumnIdx);
                }
            }else{
                if (!columnMapping.containsValue(alreadyColumnIdx)){
                    cell = dataRow.createCell(alreadyColumnIdx);
                }
                if (unmappedColumnCount.containsKey(alreadyColumnIdx)){
                    cell = dataRow.createCell(alreadyColumnIdx);
                }
            }
            if (cell != null){
                cell.setCellValue(writeConfig.getBlankValue());
                cell.setCellStyle(rowStyle);
            }
        }
    }

    @Override
    public AxolotlWriteResult finish(SXSSFSheet sheet) {
        debug(LOGGER,"结束渲染工作表[%s]",sheet.getSheetName());
        int alreadyWrittenColumns = context.getAlreadyWrittenColumns().get(context.getSwitchSheetIndex());
        // 创建结尾合计行
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING)){
            Map<Integer, BigDecimal> endingTotalMapping = context.getEndingTotalMapping().row(context.getSwitchSheetIndex());
            debug(LOGGER,"开始创建结尾合计行,合计数据为:%s",endingTotalMapping);
            SXSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
            row.setHeight((short) 600);
            for (int i = 0; i < alreadyWrittenColumns; i++) {
                SXSSFCell cell = row.createCell(i);
                String cellValue = "-";
                if (i == 0 && writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER)){
                    cellValue = "合计";
                }
                if (endingTotalMapping.containsKey(i)){
                    BigDecimal scale =
                            endingTotalMapping.get(i)
                                    .setScale(AxolotlDefaultReaderConfig.XVOID_DEFAULT_DECIMAL_SCALE, RoundingMode.HALF_UP);
                    cellValue = scale.toString();
                }
                cell.setCellValue(cellValue);
                CellStyle cellStyle = sheet.getRow(sheet.getLastRowNum() - 1).getCell(i).getCellStyle();
                SXSSFWorkbook workbook = sheet.getWorkbook();
                CellStyle totalCellStyle = workbook.createCellStyle();
                Font font = StyleHelper.createWorkBookFont(
                        workbook, fontName, true, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.BLACK
                );
                StyleHelper.setCellStyleAlignmentCenter(totalCellStyle);
                totalCellStyle.setFillForegroundColor(cellStyle.getFillForegroundColorColor());
                totalCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                totalCellStyle.setFont(font);
                BorderStyle borderStyle = BorderStyle.MEDIUM;
                totalCellStyle.setBorderBottom(borderStyle);
                totalCellStyle.setBorderLeft(borderStyle);
                totalCellStyle.setBorderRight(borderStyle);
                totalCellStyle.setBorderTop(borderStyle);
                totalCellStyle.setLeftBorderColor(cellStyle.getLeftBorderColor());
                totalCellStyle.setRightBorderColor(cellStyle.getRightBorderColor());
                totalCellStyle.setTopBorderColor(cellStyle.getTopBorderColor());
                totalCellStyle.setBottomBorderColor(cellStyle.getBottomBorderColor());
                totalCellStyle.setDataFormat(StyleHelper.DATA_FORMAT_PLAIN_TEXT_INDEX);
                cell.setCellStyle(totalCellStyle);
            }
        }
        sheet.trackAllColumnsForAutoSizing();
        for (int columnIdx = 0; columnIdx < alreadyWrittenColumns; columnIdx++) {
            sheet.autoSizeColumn(columnIdx,true);
            sheet.setColumnWidth(columnIdx, (int) (sheet.getColumnWidth(columnIdx) * 1.5));
        }
        return new AxolotlWriteResult(true, "完成结束阶段");
    }

}

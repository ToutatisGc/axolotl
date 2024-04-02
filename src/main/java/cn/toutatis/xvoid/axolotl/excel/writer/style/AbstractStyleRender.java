package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.support.*;
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
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.math.BigDecimal;
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

    /**
     * 写入配置
     */
    @Setter
    protected AutoWriteConfig writeConfig;

    /**
     * 写入上下文
     */
    @Setter
    protected AutoWriteContext context;

    /**
     * 日志
     */
    private final Logger LOGGER;

    /**
     * 字体名称
     */
    private String globalFontName;

    /**
     * 主题颜色
     */
    @Getter @Setter
    private AxolotlColor themeColor;

    /**
     * 数据写入已进行错误提示
     */
    private boolean alreadyNotice = false;

    public AbstractStyleRender(Logger LOGGER) {
        this.LOGGER = LOGGER;
    }

    public String globalFontName() {
        return globalFontName;
    }

    public void setGlobalFontName(String globalFontName) {
        this.globalFontName = globalFontName;
    }

    public static final String TOTAL_HEADER_COUNT_KEY = "";

    /**
     * 是否是第一批次数据
     * @return true/false
     */
    public boolean isFirstBatch(){
        return context.isFirstBatch(context.getSwitchSheetIndex());
    }

    /**
     * 检查并使用自定义字体
     * 如果自定义字体不为空，则使用自定义字体，否则使用默认主题字体
     * @param themeFont 主题字体
     */
    public void checkedAndUseCustomTheme(String themeFont,AxolotlColor themeColor){
        String fontName = writeConfig.getFontName();
        if (fontName != null){
            debug(LOGGER, "使用自定义字体：%s",fontName);
            setGlobalFontName(fontName);
        }else{
            setGlobalFontName(Objects.requireNonNullElse(themeFont, StyleHelper.STANDARD_FONT_NAME));
        }
        this.setThemeColor(Objects.requireNonNullElse(themeColor, StyleHelper.STANDARD_THEME_COLOR));
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
                fillWhiteCell(sheet, globalFontName);
            }
            if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH)){
                debug(LOGGER,"开启自动获取列宽");
                sheet.trackAllColumnsForAutoSizing();
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
        sheet.setDefaultRowHeight(StyleHelper.STANDARD_ROW_HEIGHT);
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
        int headerMaxDepth;
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
                row.setHeight(StyleHelper.STANDARD_ROW_HEIGHT);

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
                    headerRecursiveInfo.setRowHeight(StyleHelper.STANDARD_ROW_HEIGHT);
                    recursionRenderHeaders(sheet,childs, headerRecursiveInfo);
                }else{
                    cellAddresses = new CellRangeAddress(alreadyWriteRow, (alreadyWriteRow +headerMaxDepth)-1, headerColumnCount, headerColumnCount);

                    String fieldName = header.getFieldName();
                    if (fieldName != null){
                        debug(LOGGER,"映射字段[%s]到列索引[%s]",fieldName,headerColumnCount);
                        headerCache.put(fieldName,headerColumnCount);
                    }
                    if (!writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH)){
                        int columnWidth = header.getColumnWidth();
                        if (columnWidth < 0){
                            columnWidth = StyleHelper.getPresetCellLength(title);
                        }
                        debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",headerColumnCount,header.getName(),columnWidth);
                        sheet.setColumnWidth(headerColumnCount, columnWidth);
                    }else{
                        debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",headerColumnCount,header.getName(),"AUTO");
                    }
                    if (header.isParticipateInCalculate()){
                        debug(LOGGER,"列[%s]表头[%s]参与计算",headerColumnCount,header.getName());
                        writeConfig.addCalculateColumnIndex(headerColumnCount);
                    }
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
                        axolotlCellStyle.getFontColor(),
                        axolotlCellStyle.isItalic(),
                        axolotlCellStyle.isStrikeout()
                );
                usedCellStyle = StyleHelper.createStandardCellStyle(
                        context.getWorkbook(),
                        BorderStyle.NONE,
                        IndexedColors.BLACK,
                        axolotlCellStyle.getForegroundColor(),
                        axolotlCustomFont
                );
                usedCellStyle.setBorderTop(axolotlCellStyle.getBorderTopStyle());
                usedCellStyle.setBorderRight(axolotlCellStyle.getBorderRightStyle());
                usedCellStyle.setBorderBottom(axolotlCellStyle.getBorderBottomStyle());
                usedCellStyle.setBorderLeft(axolotlCellStyle.getBorderLeftStyle());
                usedCellStyle.setTopBorderColor(axolotlCellStyle.getTopBorderColor().getIndex());
                usedCellStyle.setRightBorderColor(axolotlCellStyle.getRightBorderColor().getIndex());
                usedCellStyle.setBottomBorderColor(axolotlCellStyle.getBottomBorderColor().getIndex());
                usedCellStyle.setLeftBorderColor(axolotlCellStyle.getLeftBorderColor().getIndex());
                usedCellStyle.setFillPattern(axolotlCellStyle.getFillPatternType());
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

                    if (!writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH)){
                        int columnWidth = header.getColumnWidth();
                        if (columnWidth == -1){
                            columnWidth = StyleHelper.getPresetCellLength(header.getName());
                        }
                        debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",alreadyWriteColumn,header.getName(),columnWidth);
                        sheet.setColumnWidth(alreadyWriteColumn, columnWidth);
                    }else{
                        debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",alreadyWriteColumn,header.getName(),"AUTO");
                    }
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
                    if (header.isParticipateInCalculate()){
                        debug(LOGGER,"列[%s]表头[%s]参与计算",alreadyWriteColumn,header.getName());
                        writeConfig.addCalculateColumnIndex(alreadyWriteColumn);
                    }
                }
            }
        }
    }

    /**
     * 属性信息
     * 用于渲染列的数据和合计
     */
    @Data
    protected static class FieldInfo{

        /**
         * 属性所属的类
         */
        private Class<?> clazz;

        /**
         * 属性名称
         */
        private final String fieldName;

        /**
         * 属性值
         */
        private final Object value;

        /**
         * 列索引
         */
        private final int columnIndex;

        /**
         * 行索引
         */
        private final int rowIndex;

        public FieldInfo(String fieldName, Object value, int columnIndex,int rowIndex) {
            if (value != null){
                this.clazz = value.getClass();
            }
            this.fieldName = fieldName;
            this.value = value;
            this.columnIndex = columnIndex;
            this.rowIndex = rowIndex;
        }
    }

    private Map<Integer, Integer> unmappedColumnCount;

    /**
     * 默认行为渲染数据
     * @param sheet 工作表
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
        dataRow.setHeight(StyleHelper.STANDARD_ROW_HEIGHT);
        int writtenColumn = START_POSITION;
        int serialNumber = context.getAndIncrementSerialNumber() - context.getHeaderRowCount().get(switchSheetIndex);
        // 写入数据
        Map<String, Integer> columnMapping = context.getHeaderColumnIndexMapping().row(context.getSwitchSheetIndex());
        unmappedColumnCount =  new HashMap<>();
        columnMapping.forEach((key, value) -> unmappedColumnCount.put(value, 1));
        boolean columnMappingEmpty = columnMapping.isEmpty();
        // 写入序号
        boolean autoInsertSerialNumber = writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER);
        if (autoInsertSerialNumber){
            SXSSFCell cell = dataRow.createCell(writtenColumn);
            cell.setCellValue(serialNumber);
            cell.setCellStyle(rowStyle);
            writtenColumnMap.put(columnMappingEmpty?writtenColumn++:0,1);
        }
        for (Map.Entry<String, Object> dataEntry : dataMap.entrySet()) {
            String fieldName = dataEntry.getKey();
            SXSSFCell cell;
            if (columnMappingEmpty){
                cell = dataRow.createCell(writtenColumn);
            }else{
                if (columnMapping.containsKey(fieldName)){
                    cell = (SXSSFCell) ExcelToolkit.createOrCatchCell(sheet,alreadyWriteRow,columnMapping.get(fieldName),null);
                }else {
                    if (!alreadyNotice){
                        warn(LOGGER,"未映射字段[%s]请在表头Header中映射字段!",fieldName);
                        alreadyNotice = true;
                    }
                    continue;
                }
            }
            Object value = dataEntry.getValue();
            int columnNumber = columnMappingEmpty ? writtenColumn : columnMapping.get(fieldName);
            FieldInfo fieldInfo = new FieldInfo(fieldName, value,columnNumber ,alreadyWriteRow);
            cell.setCellStyle(rowStyle);
            // 渲染数据到单元格
            this.renderColumn(fieldInfo,cell);
            if (columnMappingEmpty){
                writtenColumnMap.put(writtenColumn++,1);
            }else{
                writtenColumnMap.put(columnNumber,1);
            }
        }
        // 将未使用的的单元格赋予空值
        for (int alreadyColumnIdx = 0; alreadyColumnIdx < context.getAlreadyWrittenColumns().get(switchSheetIndex); alreadyColumnIdx++) {
            if (autoInsertSerialNumber && alreadyColumnIdx == 0){
                continue;
            }
            SXSSFCell cell = null;
            if (columnMappingEmpty){
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

    /**
     * 渲染列数据
     * @param fieldInfo 字段信息
     * @param cell 单元格
     */
    public void renderColumn(FieldInfo fieldInfo,Cell cell){
        Object value = fieldInfo.getValue();
        if (value == null){
            cell.setCellValue(writeConfig.getBlankValue());
        }else{
            calculateColumns(fieldInfo);
            value = writeConfig.getDataInverter().convert(value);
            int columnIndex = fieldInfo.getColumnIndex();
            cell.setCellValue(value.toString());
            unmappedColumnCount.remove(columnIndex);
        }
    }

    /**
     * 计算列合计
     * @param fieldInfo 字段信息
     */
    public void calculateColumns(FieldInfo fieldInfo){
        int columnIndex = fieldInfo.getColumnIndex();
        String value = fieldInfo.getValue().toString();
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING) && Validator.strIsNumber(value)){
            Map<Integer, BigDecimal> endingTotalMapping = context.getEndingTotalMapping().row(context.getSwitchSheetIndex());
            if (endingTotalMapping.containsKey(columnIndex)){
                BigDecimal newValue = endingTotalMapping.get(columnIndex).add(BigDecimal.valueOf(Double.parseDouble(value)));
                endingTotalMapping.put(columnIndex,newValue);
            }else{
                endingTotalMapping.put(columnIndex,BigDecimal.valueOf(Double.parseDouble(value)));
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
            DataInverter<?> dataInverter = writeConfig.getDataInverter();
            HashSet<Integer> calculateColumnIndexes = writeConfig.getCalculateColumnIndexes();
            for (int i = 0; i < alreadyWrittenColumns; i++) {
                SXSSFCell cell = row.createCell(i);
                String cellValue = "-";
                if (i == 0 && writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER)){
                    cellValue = "合计";
                    sheet.setColumnWidth(i, StyleHelper.SERIAL_NUMBER_LENGTH);
                }
                if (endingTotalMapping.containsKey(i)){
                    if (calculateColumnIndexes.contains(i) || (calculateColumnIndexes.size() == 1 && calculateColumnIndexes.contains(-1))){
                        BigDecimal bigDecimal = endingTotalMapping.get(i);
                        Object convert = dataInverter.convert(bigDecimal);
                        cellValue = convert.toString();
                    }
                }
                cell.setCellValue(cellValue);
                CellStyle cellStyle = sheet.getRow(sheet.getLastRowNum() - 1).getCell(i).getCellStyle();
                SXSSFWorkbook workbook = sheet.getWorkbook();
                CellStyle totalCellStyle = workbook.createCellStyle();
                Font font = StyleHelper.createWorkBookFont(
                        workbook, globalFontName, true, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.BLACK
                );
                font.setColor(workbook.getFontAt(cellStyle.getFontIndex()).getColor());
                StyleHelper.setCellStyleAlignmentCenter(totalCellStyle);
                totalCellStyle.setFillForegroundColor(cellStyle.getFillForegroundColorColor());
                totalCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                totalCellStyle.setFont(font);
                BorderStyle borderStyle = BorderStyle.THIN;
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
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH)){
            debug(LOGGER,"开始自动计算列宽");
            for (int columnIdx = 0; columnIdx < alreadyWrittenColumns; columnIdx++) {
                sheet.autoSizeColumn(columnIdx,true);
                sheet.setColumnWidth(columnIdx, (int) (sheet.getColumnWidth(columnIdx) * 1.35));
            }
        }
        Map<Integer, Integer> specialRowHeightMapping = writeConfig.getSpecialRowHeightMapping();
        for (Map.Entry<Integer, Integer> heightEntry : specialRowHeightMapping.entrySet()) {
            debug(LOGGER,"设置工作表[%s]第%s行高度为%s",sheet.getSheetName(),heightEntry.getKey(),heightEntry.getValue());
            sheet.getRow(heightEntry.getKey()).setHeightInPoints(heightEntry.getValue());
        }
        return new AxolotlWriteResult(true, "完成结束阶段");
    }

    /* 拓展方法部分 */

    /**
     * 创建字体
     * @param fontName 字体名称
     * @param fontSize 字体大小
     * @param isBold 是否加粗
     * @param color 颜色
     * @return 字体
     */
    public Font createFont(String fontName,short fontSize,boolean isBold,IndexedColors color){
        return StyleHelper.createWorkBookFont(context.getWorkbook(),fontName,isBold, fontSize,color);
    }

    /**
     * 创建字体
     * @param fontName 字体名称
     * @param fontSize 字体大小
     * @param isBold 是否加粗
     * @param color 颜色
     * @return 字体
     */
    public Font createFont(String fontName,short fontSize,boolean isBold,AxolotlColor color){
        XSSFFont font = new XSSFFont();
        font.setColor(color.toXSSFColor());
        font.setBold(isBold);
        font.setFontName(fontName);
        font.setFontHeightInPoints(fontSize);
        StylesTable stylesSource = context.getWorkbook().getXSSFWorkbook().getStylesSource();
        font.registerTo(stylesSource);
        return font;
    }

    /**
     * 创建字体
     * @param fontName 字体名称
     * @param fontSize 字体大小
     * @param isBold 是否加粗
     * @param color 颜色
     * @param italic 是否斜体
     * @param strikeout 是否删除线
     * @return
     */
    public Font createFont(String fontName,short fontSize,boolean isBold,IndexedColors color,boolean italic,boolean strikeout){
        return StyleHelper.createWorkBookFont(context.getWorkbook(),fontName,isBold, fontSize,color,italic,strikeout);
    }

    /**
     * 创建字体
     * @param fontName 字体名称
     * @param fontSize 字体大小
     * @param isBold 是否加粗
     * @param color 颜色
     * @param italic 是否斜体
     * @param strikeout 是否删除线
     * @return
     */
    public Font createFont(String fontName,short fontSize,boolean isBold,AxolotlColor color,boolean italic,boolean strikeout){
        XSSFFont font = new XSSFFont();
        font.setColor(color.toXSSFColor());
        font.setBold(isBold);
        font.setFontName(fontName);
        font.setFontHeightInPoints(fontSize);
        font.setItalic(italic);
        font.setStrikeout(strikeout);
        StylesTable stylesSource = context.getWorkbook().getXSSFWorkbook().getStylesSource();
        font.registerTo(stylesSource);
        return font;
    }

    public Font createMainTextFont(short fontSize,AxolotlColor color){return this.createFont(globalFontName, fontSize, false, color);}
    public Font createMainTextFont(AxolotlColor color){return this.createMainTextFont(StyleHelper.STANDARD_TEXT_FONT_SIZE, color);}
    public Font createMainTextFont(short fontSize,IndexedColors color){return this.createFont(globalFontName, fontSize, false, color);}
    public Font createMainTextFont(IndexedColors color){return this.createMainTextFont(StyleHelper.STANDARD_TEXT_FONT_SIZE, color);}
    public Font createBlackMainTextFont(){return this.createMainTextFont(IndexedColors.BLACK);}
    public Font createWhiteMainTextFont(){return this.createMainTextFont(IndexedColors.WHITE);}
    public Font createRedMainTextFont(){return this.createMainTextFont(IndexedColors.RED);}

    public CellStyle createBlackMainTextCellStyle(IndexedColors borderColor, AxolotlColor cellColor){
        return createStyle(BorderStyle.THIN, borderColor, cellColor,
                globalFontName, StyleHelper.STANDARD_TEXT_FONT_SIZE, false, IndexedColors.BLACK);
    }

    public CellStyle createBlackMainTextCellStyle(BorderStyle borderStyle, IndexedColors borderColor, AxolotlColor cellColor){
        return createStyle(borderStyle, borderColor, cellColor,
                globalFontName, StyleHelper.STANDARD_TEXT_FONT_SIZE, false, IndexedColors.BLACK);
    }

    public CellStyle createWhiteMainTextCellStyle(BorderStyle borderStyle, IndexedColors borderColor, AxolotlColor cellColor){
        return createStyle(borderStyle, borderColor, cellColor,
                globalFontName, StyleHelper.STANDARD_TEXT_FONT_SIZE, false, IndexedColors.WHITE);
    }

    public CellStyle createStyle(BorderStyle borderStyle, IndexedColors borderColor, AxolotlColor cellColor,Font font) {
        return StyleHelper.createStandardCellStyle(context.getWorkbook(),borderStyle, borderColor,cellColor,font);
    }

    public CellStyle createStyle(BorderStyle borderStyle, IndexedColors borderColor, AxolotlColor cellColor,
                                 String fontName,short fontSize,boolean isBold,Object fontColor) {
        Font font;
        if (fontColor instanceof AxolotlColor){
            font = this.createFont(fontName,fontSize,isBold, (AxolotlColor) fontColor);
        }else if (fontColor instanceof IndexedColors){
            font = this.createFont(fontName,fontSize,isBold, (IndexedColors) fontColor);
        }else{
            throw new IllegalArgumentException("字体颜色类型错误");
        }
        return StyleHelper.createStandardCellStyle(context.getWorkbook(),borderStyle, borderColor,cellColor,font);
    }

    public CellStyle createStyle(BorderStyle borderStyle, IndexedColors borderColor, AxolotlColor cellColor,
                                 String fontName,short fontSize,boolean isBold,Object fontColor,boolean italic,boolean strikeout) {
        Font font;
        if (fontColor instanceof AxolotlColor){
            font = this.createFont(fontName,fontSize,isBold,(AxolotlColor) fontColor,italic,strikeout);
        }else if (fontColor instanceof IndexedColors){
            font = this.createFont(fontName,fontSize,isBold,(IndexedColors) fontColor, italic,strikeout);
        }else{
            throw new IllegalArgumentException("字体颜色类型错误");
        }
        return StyleHelper.createStandardCellStyle(context.getWorkbook(),borderStyle, borderColor,cellColor,font);
    }

}

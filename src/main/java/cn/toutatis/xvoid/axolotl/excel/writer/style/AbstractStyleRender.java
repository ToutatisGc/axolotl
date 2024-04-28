package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.Meta;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.annotations.AxolotlWriteIgnore;
import cn.toutatis.xvoid.axolotl.excel.writer.components.annotations.AxolotlWriterGetter;
import cn.toutatis.xvoid.axolotl.excel.writer.components.configuration.AxolotlCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.configuration.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.AutoWriteContext;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.support.inverters.DataInverter;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Data;
import lombok.Getter;
import lombok.Setter;
import lombok.SneakyThrows;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.SpreadsheetVersion;
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
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
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
     * 组件渲染器
     */
    @Getter @Setter
    private ComponentRender componentRender = new ComponentRender() {};

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
                int sheetIndex = context.getSwitchSheetIndex();
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
            int sheetIndex = context.getSwitchSheetIndex();
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
                        writeConfig.addCalculateColumnIndex(sheetIndex,headerColumnCount);
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
            return header.getCustomCellStyle();
        }else{
            AxolotlCellStyle axolotlCellStyle = header.getAxolotlCellStyle();
            if (axolotlCellStyle != null){
                SXSSFWorkbook workbook = context.getWorkbook();
                CellStyle cellStyle = workbook.createCellStyle();
                //用默认样式给新样式赋值
                cellStyle.setBorderTop(usedCellStyle.getBorderTop());
                cellStyle.setBorderRight(usedCellStyle.getBorderRight());
                cellStyle.setBorderBottom(usedCellStyle.getBorderBottom());
                cellStyle.setBorderLeft(usedCellStyle.getBorderLeft());
                cellStyle.setTopBorderColor(usedCellStyle.getTopBorderColor());
                cellStyle.setRightBorderColor(usedCellStyle.getRightBorderColor());
                cellStyle.setBottomBorderColor(usedCellStyle.getBottomBorderColor());
                cellStyle.setLeftBorderColor(usedCellStyle.getLeftBorderColor());
                cellStyle.setFillPattern(usedCellStyle.getFillPattern());
                cellStyle.setDataFormat(usedCellStyle.getDataFormat());
                cellStyle.setFillForegroundColor(usedCellStyle.getFillForegroundColorColor());
                cellStyle.setAlignment(usedCellStyle.getAlignment());
                cellStyle.setVerticalAlignment(usedCellStyle.getVerticalAlignment());
                //根据配置修改新样式的值
                if(axolotlCellStyle.getBorderLeftStyle() != null){
                    cellStyle.setBorderLeft(axolotlCellStyle.getBorderLeftStyle());
                }
                if(axolotlCellStyle.getBorderRightStyle() != null){
                    cellStyle.setBorderRight(axolotlCellStyle.getBorderRightStyle());
                }
                if(axolotlCellStyle.getBorderTopStyle() != null){
                    cellStyle.setBorderTop(axolotlCellStyle.getBorderTopStyle());
                }
                if(axolotlCellStyle.getBorderBottomStyle() != null){
                    cellStyle.setBorderBottom(axolotlCellStyle.getBorderBottomStyle());
                }
                if(axolotlCellStyle.getLeftBorderColor() != null){
                    cellStyle.setLeftBorderColor(axolotlCellStyle.getLeftBorderColor().getIndex());
                }
                if(axolotlCellStyle.getRightBorderColor() != null){
                    cellStyle.setRightBorderColor(axolotlCellStyle.getRightBorderColor().getIndex());
                }
                if(axolotlCellStyle.getTopBorderColor() != null){
                    cellStyle.setTopBorderColor(axolotlCellStyle.getTopBorderColor().getIndex());
                }
                if(axolotlCellStyle.getBottomBorderColor() != null){
                    cellStyle.setBottomBorderColor(axolotlCellStyle.getBottomBorderColor().getIndex());
                }
                if(axolotlCellStyle.getForegroundColor() != null){
                    cellStyle.setFillForegroundColor(axolotlCellStyle.getForegroundColor());
                }
                if(axolotlCellStyle.getFillPatternType() != null){
                    cellStyle.setFillPattern(axolotlCellStyle.getFillPatternType());
                }
                if(axolotlCellStyle.getHorizontalAlignment() != null){
                    cellStyle.setAlignment(axolotlCellStyle.getHorizontalAlignment());
                }
                if(axolotlCellStyle.getVerticalAlignment() != null){
                    cellStyle.setVerticalAlignment(axolotlCellStyle.getVerticalAlignment());
                }

                //获取默认字体
                Font fontAt = workbook.getFontAt(usedCellStyle.getFontIndex());

                //创建新字体 用默认字体赋值
                Font font = workbook.createFont();
                font.setFontName(fontAt.getFontName());
                font.setBold(fontAt.getBold());
                font.setFontHeightInPoints(fontAt.getFontHeightInPoints());
                font.setItalic(fontAt.getItalic());
                font.setStrikeout(fontAt.getStrikeout());
                font.setColor(fontAt.getColor());

                //根据配置修改新字体的值
                if(axolotlCellStyle.getFontName() != null){
                    font.setFontName(axolotlCellStyle.getFontName());
                }
                if(axolotlCellStyle.getFontSize() != null){
                    font.setFontHeightInPoints(axolotlCellStyle.getFontSize());
                }
                if(axolotlCellStyle.getFontColor() != null){
                    font.setColor(axolotlCellStyle.getFontColor().getIndex());
                }
                if(axolotlCellStyle.getFontBold() != null){
                    font.setBold(axolotlCellStyle.getFontBold());
                }
                if(axolotlCellStyle.getItalic() != null){
                    font.setItalic(axolotlCellStyle.getItalic());
                }
                if(axolotlCellStyle.getStrikeout() != null){
                    font.setStrikeout(axolotlCellStyle.getStrikeout());
                }
                cellStyle.setFont(font);
                return cellStyle;
            }else{
                return usedCellStyle;
            }
        }
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
            int sheetIndex = context.getSwitchSheetIndex();
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
                        writeConfig.addCalculateColumnIndex(sheetIndex,alreadyWriteColumn);
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
    public static class FieldInfo{

        /**
         * 属性类型
         */
        private Class<?> clazz;

        /**
         * 数据实例
         */
        private final Object dataInstance;

        /**
         * 属性名称
         */
        private final String fieldName;

        /**
         * 属性值
         */
        private final Object value;

        private final int sheetIndex;

        /**
         * 列索引
         */
        private final int columnIndex;

        /**
         * 行索引
         */
        private final int rowIndex;

        public FieldInfo(Object dataInstance, String fieldName, Object value,int sheetIndex, int columnIndex,int rowIndex) {
            this.dataInstance = dataInstance;
            if (value != null){
                this.clazz = value.getClass();
            }
            this.fieldName = fieldName;
            this.value = value;
            this.columnIndex = columnIndex;
            this.rowIndex = rowIndex;
            this.sheetIndex = sheetIndex;
        }
    }

    protected Map<Integer, Integer> unmappedColumnCount;

    /**
     * 默认行为渲染数据
     * @param sheet 工作表
     */
    public void defaultRenderNextData(SXSSFSheet sheet,Object data,CellStyle rowStyle){
        // 获取对象属性
        HashMap<String, Object> dataMap = getDataMap(data);
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
            FieldInfo fieldInfo = new FieldInfo(data, fieldName, value, switchSheetIndex, columnNumber, alreadyWriteRow);
            cell.setCellStyle(rowStyle);
            // 渲染数据到单元格
            this.renderColumn(fieldInfo,cell);
            unmappedColumnCount.remove(fieldInfo.getColumnIndex());
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
     * 获取对象属性
     * @param data 数据对象
     * @return 属性集合
     */
    @SuppressWarnings({"rawtypes","unchecked"})
    public HashMap<String, Object> getDataMap(Object data){
        if(data == null){return new LinkedHashMap<>();}
        // 获取对象属性
        HashMap<String, Object> dataMap = new LinkedHashMap<>();
        if (data instanceof Map map) {
            map.keySet().forEach(key -> {
                if (!key.toString().startsWith(Meta.MODULE_NAME.toUpperCase())){
                    dataMap.put(key.toString(), map.get(key));
                }
            });
        }else{
            Class<?> dataClass = data.getClass();
            if(writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_USE_GETTER_METHOD)){
                ArrayList<Method> getterMethods = ReflectToolkit.getGetterMethods(dataClass);
                Method tmpMethod;
                for (Method getterMethod : getterMethods) {
                    // 仅获取公开的Getter方法
                    if (Modifier.isPublic(getterMethod.getModifiers())){
                        tmpMethod = getterMethod;
                        AxolotlWriterGetter axolotlWriterGetter = tmpMethod.getAnnotation(AxolotlWriterGetter.class);
                        if (axolotlWriterGetter != null && axolotlWriterGetter.value() != null){
                            try {
                                debug(LOGGER,"Getter方法[%s]将被重定向到[%s]",tmpMethod.getName(),axolotlWriterGetter.value());
                                tmpMethod = dataClass.getMethod(axolotlWriterGetter.value());
                                if (!Modifier.isPublic(getterMethod.getModifiers())){
                                    String message = format("重定向Getter方法[%s]失败,方法[%s]为私有方法", tmpMethod.getName(), axolotlWriterGetter.value());
                                    error(LOGGER,message);
                                    throw new AxolotlWriteException(message);
                                }
                            } catch (NoSuchMethodException e) {
                                String message = format("Getter方法[%s]重定向失败,方法不存在", tmpMethod.getName());
                                error(LOGGER,message);
                                throw new AxolotlWriteException(message);
                            }
                        }
                        AxolotlWriteIgnore ignore = getterMethod.getAnnotation(AxolotlWriteIgnore.class);
                        if (ignore != null){continue;}
                        var methodName = getterMethod.getName();
                        String fieldName = ReflectToolkit.convertGetterToFieldName(methodName);
                        Field field = ReflectToolkit.recursionGetField(dataClass, fieldName);
                        if (field != null){
                            ignore = field.getAnnotation(AxolotlWriteIgnore.class);
                            if (ignore != null){continue;}
                        }
                        int parameterCount = tmpMethod.getParameterCount();
                        if (parameterCount == 0){
                            Object invokeValue = null;
                            try {
                                invokeValue = tmpMethod.invoke(data);
                            } catch (IllegalAccessException | InvocationTargetException e) {
                                e.printStackTrace();
                                if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.SIMPLE_EXCEPTION_RETURN_RESULT)){
                                    LoggerHelper.error(LOGGER,format("[%s]方法调用失败,将赋予null值",methodName));
                                }else{
                                    throw new AxolotlWriteException(e.getMessage());
                                }
                            }
                            dataMap.put(fieldName,invokeValue);
                        }else{
                            LoggerHelper.debug(LOGGER,format("方法[%s]参数数量大于0将跳过",getterMethod.getName()));
                        }
                    }else{
                        LoggerHelper.debug(LOGGER,format("方法[%s]为私有方法将跳过",getterMethod.getName()));
                    }
                }
            }else {
                List<Field> fields = ReflectToolkit.getAllFields(dataClass, true);
                fields.forEach(field -> {
                    field.setAccessible(true);
                    String fieldName = field.getName();
                    try {
                        dataMap.put(fieldName, field.get(data));
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                        throw new AxolotlWriteException("获取对象字段错误");
                    }
                });
            }
        }
        return dataMap;
    }

    /**
     * 渲染列信息
     * @param fieldInfo 字段信息
     * @param cell 单元格
     */
    public void renderColumn(FieldInfo fieldInfo,Cell cell){
        Object value = fieldInfo.getValue();
        if (value == null){
            componentRender.renderFieldColumnNullValue(fieldInfo,cell);
        }else {
            componentRender.defaultRenderColumn(fieldInfo,cell);
        }
    }


    @Override
    public AxolotlWriteResult finish(SXSSFSheet sheet) {
        debug(LOGGER,"结束渲染工作表[%s]",sheet.getSheetName());
        int sheetIndex = context.getWorkbook().getSheetIndex(sheet);
        int alreadyWrittenColumns = context.getAlreadyWrittenColumns().get(sheetIndex);
        // 创建结尾合计行
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING)){
            Map<Integer, BigDecimal> endingTotalMapping = context.getEndingTotalMapping().row(sheetIndex);
            debug(LOGGER,"开始创建结尾合计行,合计数据为:%s",endingTotalMapping);
            SXSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
            row.setHeight((short) 600);
            DataInverter<?> dataInverter = writeConfig.getDataInverter();
            Set<Integer> calculateColumnIndexes = writeConfig.getCalculateColumnIndexes(sheetIndex);
            for (int i = 0; i < alreadyWrittenColumns; i++) {
                SXSSFCell cell = row.createCell(i);
                String cellValue = writeConfig.getBlankValue();
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
        Map<Integer, Integer> specialRowHeightMapping = writeConfig.getSpecialRowHeightMapping(sheetIndex);
        for (Map.Entry<Integer, Integer> heightEntry : specialRowHeightMapping.entrySet()) {
            debug(LOGGER,"设置工作表[%s]第%s行高度为%s",sheet.getSheetName(),heightEntry.getKey(),heightEntry.getValue());
            sheet.getRow(heightEntry.getKey()).setHeightInPoints(heightEntry.getValue());
        }
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_HIDDEN_BLANK_COLUMNS)){
            warn(LOGGER,"即将隐藏空白列，增加导出耗时");
            int maxColumnIndex = -1;
            for (Row cells : sheet) {
                maxColumnIndex = Math.max(cells.getLastCellNum(),maxColumnIndex);
            }
            if (maxColumnIndex > -1){
                for (int i = maxColumnIndex; i < SpreadsheetVersion.EXCEL2007.getMaxColumns(); i++) {
                    sheet.setColumnHidden(i,true);
                }
            }
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

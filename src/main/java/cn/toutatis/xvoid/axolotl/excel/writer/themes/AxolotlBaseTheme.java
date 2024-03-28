package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.BaseCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.CellStyleConfigur;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.SneakyThrows;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.checkerframework.checker.units.qual.A;
import org.slf4j.Logger;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper.START_POSITION;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.debug;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.info;

/**
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/28 9:21
 */
public class AxolotlBaseTheme extends AbstractStyleRender implements ExcelStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlAdministrationRedTheme.class);

    private BaseCellStyle globalCellStyle;

    private CellStyleConfigur cellStyleConfigur;

    public AxolotlBaseTheme() {
        super(LOGGER);
        globalCellStyle = getDefaultCellStyle();
    }

    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        AxolotlWriteResult axolotlWriteResult;
        if(isFirstBatch()){
            BaseCellStyle cellStyle = cellStyleConfigur.globalCellStyle();
            if(cellStyle != null){
                addConfig(cellStyle,globalCellStyle);
            }
            String fontName = writeConfig.getFontName();
            if (fontName != null){
                debug(LOGGER, "使用自定义字体：%s",fontName);
                globalCellStyle.setFontName(fontName);
            }

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
                this.fillWhiteCell(sheet);
            }
        }else {
            axolotlWriteResult = new AxolotlWriteResult(true,"已初始化");
        }
        return axolotlWriteResult;
    }

    @Override
    public AxolotlWriteResult renderHeader(SXSSFSheet sheet) {
        // 1.创建标题行
        AxolotlWriteResult writeTitle = createTitleRow(sheet);

        // 2.创建表头单元格样式
        XSSFCellStyle headerDefaultCellStyle = (XSSFCellStyle) createStyle(globalCellStyle.getBaseBorderStyle(), globalCellStyle.getBaseBorderColor(), globalCellStyle.getForegroundColor(), globalCellStyle.getFontName(), globalCellStyle.getFontSize(), globalCellStyle.isBold(), globalCellStyle.getFontColor());
        StyleHelper.setCellStyleAlignment(headerDefaultCellStyle, globalCellStyle.getHorizontalAlignment(), globalCellStyle.getVerticalAlignment());
        setBorderStyle(headerDefaultCellStyle,globalCellStyle);
        headerDefaultCellStyle.setWrapText(true);
        headerDefaultCellStyle.setFillPattern(globalCellStyle.getFillPatternType());

        // 3.渲染表头
        AxolotlWriteResult headerWriteResult = this.defaultRenderHeaders(sheet, headerDefaultCellStyle);

        // 4.合并表头
        if (writeTitle.isWrite()){
            this.mergeTitleRegion(sheet,context.getAlreadyWrittenColumns().get(context.getSwitchSheetIndex()),headerDefaultCellStyle);
        }

        // 5.创建冻结窗格
        sheet.createFreezePane(START_POSITION, context.getAlreadyWriteRow().get(context.getSwitchSheetIndex())+1);

        return headerWriteResult;
    }

    @Override
    public AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data) {
        return null;
    }


    /**
     * 创建预制样式
     * @return
     */
    private BaseCellStyle getDefaultCellStyle(){
        BaseCellStyle baseCellStyle = new BaseCellStyle();
        baseCellStyle.setTitleRowHeight(StyleHelper.STANDARD_TITLE_ROW_HEIGHT);
        baseCellStyle.setHeaderRowHeight(StyleHelper.STANDARD_HEADER_ROW_HEIGHT);
        baseCellStyle.setDataRowHeight(StyleHelper.STANDARD_HEADER_ROW_HEIGHT);

        baseCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        baseCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        baseCellStyle.setForegroundColor(new AxolotlColor(255,255,255));
        baseCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);

        baseCellStyle.setBaseBorderStyle(BorderStyle.NONE);
        baseCellStyle.setBaseBorderColor(IndexedColors.BLACK);
        baseCellStyle.setFontName(StyleHelper.STANDARD_FONT_NAME);
        baseCellStyle.setFontColor(IndexedColors.BLACK);
        baseCellStyle.setBold(false);
        baseCellStyle.setFontSize(StyleHelper.STANDARD_TEXT_FONT_SIZE);

        return baseCellStyle;
    }

    private void addConfig(BaseCellStyle cellStyle,BaseCellStyle defaultCellStyle){
        if(!cellStyle.getClass().equals(BaseCellStyle.class)){
            return;
        }
        List<Field> fields = ReflectToolkit.getAllFields(BaseCellStyle.class, false);
        for (Field field : fields) {
            try {
                field.setAccessible(true);
                Object v = field.get(cellStyle);
                if(v != null){
                    ReflectToolkit.setObjectField(defaultCellStyle,field,v);
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 填充空白单元格
     * @param sheet 工作表
     */
    public void fillWhiteCell(Sheet sheet){
       // Font font = createFont(globalCellStyle.getFontName(), globalCellStyle.getFontSize(), globalCellStyle.isBold(), globalCellStyle.getFontColor());
        CellStyle defaultStyle = createStyle(globalCellStyle.getBaseBorderStyle(), globalCellStyle.getBaseBorderColor(), globalCellStyle.getForegroundColor(), globalCellStyle.getFontName(), globalCellStyle.getFontSize(), globalCellStyle.isBold(), globalCellStyle.getFontColor());
        //设置单元格对齐方式
        StyleHelper.setCellStyleAlignment(defaultStyle, globalCellStyle.getHorizontalAlignment(), globalCellStyle.getVerticalAlignment());
        //设置单元格边框样式
        setBorderStyle(defaultStyle,globalCellStyle);
        //填充样式
        defaultStyle.setFillPattern(globalCellStyle.getFillPatternType());
        // 将默认样式应用到所有单元格
        for (int i = 0; i < 26; i++) {
            sheet.setDefaultColumnStyle(i, defaultStyle);
            sheet.setDefaultColumnWidth(12);
        }
        sheet.setDefaultRowHeight((short) 400);
    }

    /**
     * 根据全局配置设置单元格边框样式
     * @param cellStyle 单元格样式
     * @param baseCellStyle 全局配置
     */
    private void setBorderStyle(CellStyle cellStyle,BaseCellStyle baseCellStyle){
        BorderStyle borderTopStyle = baseCellStyle.getBorderTopStyle();
        if(borderTopStyle != null){
            cellStyle.setBorderTop(borderTopStyle);
        }
        IndexedColors topBorderColor = baseCellStyle.getTopBorderColor();
        if(topBorderColor != null){
            cellStyle.setTopBorderColor(topBorderColor.getIndex());
        }
        BorderStyle borderBottomStyle = baseCellStyle.getBorderBottomStyle();
        if(borderBottomStyle != null){
            cellStyle.setBorderBottom(borderBottomStyle);
        }
        IndexedColors bottomBorderColor = baseCellStyle.getBottomBorderColor();
        if(bottomBorderColor != null){
            cellStyle.setBottomBorderColor(bottomBorderColor.getIndex());
        }
        BorderStyle borderLeftStyle = baseCellStyle.getBorderLeftStyle();
        if(borderLeftStyle != null){
            cellStyle.setBorderLeft(borderLeftStyle);
        }
        IndexedColors leftBorderColor = baseCellStyle.getLeftBorderColor();
        if(leftBorderColor != null){
            cellStyle.setLeftBorderColor(leftBorderColor.getIndex());
        }
        BorderStyle borderRightStyle = baseCellStyle.getBorderRightStyle();
        if(borderRightStyle != null){
            cellStyle.setBorderRight(borderRightStyle);
        }
        IndexedColors rightBorderColor = baseCellStyle.getRightBorderColor();
        if(rightBorderColor != null){
            cellStyle.setRightBorderColor(rightBorderColor.getIndex());
        }
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
            titleRow.setHeight(globalCellStyle.getTitleRowHeight());
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
     * Part.1 表头
     * Step.2 递归表头
     * @param sheet 工作表
     * @param headerDefaultCellStyle 表头默认样式
     */
    public AxolotlWriteResult defaultRenderHeaders(SXSSFSheet sheet, CellStyle headerDefaultCellStyle){
        int switchSheetIndex = context.getSwitchSheetIndex();
        List<cn.toutatis.xvoid.axolotl.excel.writer.components.Header> headers = context.getHeaders().get(switchSheetIndex);
        int headerMaxDepth;
        int headerColumnCount = 0;
        int alreadyWriteRow = context.getAlreadyWriteRow().getOrDefault(context.getSwitchSheetIndex(),-1);
        if (headers != null && !headers.isEmpty()){
            List<cn.toutatis.xvoid.axolotl.excel.writer.components.Header> cacheHeaders;
            if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER)){
                cacheHeaders = new ArrayList<>();
                cacheHeaders.add(new cn.toutatis.xvoid.axolotl.excel.writer.components.Header("序号"));
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
                row.setHeight(globalCellStyle.getHeaderRowHeight());

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
                    headerRecursiveInfo.setRowHeight(globalCellStyle.getHeaderRowHeight());
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
                if(axolotlCellStyle.getBorderStyle() != null){
                    usedCellStyle.setBorderTop(axolotlCellStyle.getBorderStyle());
                    usedCellStyle.setBorderRight(axolotlCellStyle.getBorderStyle());
                    usedCellStyle.setBorderBottom(axolotlCellStyle.getBorderStyle());
                    usedCellStyle.setBorderLeft(axolotlCellStyle.getBorderStyle());
                }
                if(axolotlCellStyle.getBorderColor() != null){
                    usedCellStyle.setTopBorderColor(axolotlCellStyle.getBorderColor().getIndex());
                    usedCellStyle.setRightBorderColor(axolotlCellStyle.getBorderColor().getIndex());
                    usedCellStyle.setBottomBorderColor(axolotlCellStyle.getBorderColor().getIndex());
                    usedCellStyle.setLeftBorderColor(axolotlCellStyle.getBorderColor().getIndex());
                }
                if(axolotlCellStyle.getForegroundColor() != null){
                    usedCellStyle.setFillBackgroundColor(axolotlCellStyle.getForegroundColor());
                }
                if(axolotlCellStyle.getFillPatternType() != null){
                    usedCellStyle.setFillPattern(axolotlCellStyle.getFillPatternType());
                }

                Font font = createFont(globalCellStyle.getFontName(), globalCellStyle.getFontSize(), globalCellStyle.isBold(), globalCellStyle.getFontColor());
                if(axolotlCellStyle.getFontName() != null){
                    font.setFontName(axolotlCellStyle.getFontName());
                }
                if(axolotlCellStyle.getFontSize() != -1){
                    font.setFontHeightInPoints(axolotlCellStyle.getFontSize());
                }
                if(axolotlCellStyle.getFontColor() != null){
                    font.setColor(axolotlCellStyle.getFontColor().getIndex());
                }
                font.setBold(axolotlCellStyle.isFontBold());

                usedCellStyle.setFont(font);
            }
        }
        return usedCellStyle;
    }


}

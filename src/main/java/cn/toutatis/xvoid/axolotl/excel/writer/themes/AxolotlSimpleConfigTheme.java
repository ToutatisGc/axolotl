package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.*;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.CellStyleConfigur;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.DataInverter;
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
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.slf4j.Logger;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.*;

import static cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper.START_POSITION;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

/**
 * 简单配置 主题
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/28 9:21
 */
public class AxolotlSimpleConfigTheme extends AbstractStyleRender implements ExcelStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlAdministrationRedTheme.class);

    /**
     * 全局样式
     */
    private GlobalCellStyle globalCellStyle;

    /**
     * 表头样式
     */
    private GlobalCellStyle handlerCellStyle;

    /**
     * 标题样式
     */
    private GlobalCellStyle titleCellStyle;

    /**
     * 程序常用样式
     */
    private GlobalCellStyle commonCellStyle;

    /**
     * 样式创建缓存
     */
    private Map<GlobalCellStyle,CellStyle> cellStyleCache = new HashMap<>();

    private Short titleRowHeight = StyleHelper.STANDARD_TITLE_ROW_HEIGHT;

    private Short headerRowHeight = StyleHelper.STANDARD_HEADER_ROW_HEIGHT;

    private Short dataRowHeight = StyleHelper.STANDARD_HEADER_ROW_HEIGHT;

    private Short commonRowHeight = StyleHelper.STANDARD_HEADER_ROW_HEIGHT;

    private Short columnWidth = (short) 12;

    /**
     * 样式配置类
     */
    private CellStyleConfigur cellStyleConfigur;

    public AxolotlSimpleConfigTheme() {
        super(LOGGER);
        throw new AxolotlWriteException("样式配置类初始化失败，请使用有参构造进行创建");
    }

    public AxolotlSimpleConfigTheme(CellStyleConfigur cellStyleConfigur) {
        super(LOGGER);
        this.cellStyleConfigur = cellStyleConfigur;
    }

    public AxolotlSimpleConfigTheme(Class<? extends CellStyleConfigur> configurClass) {
        super(LOGGER);
        try {
            this.cellStyleConfigur = configurClass.getDeclaredConstructor().newInstance();
        } catch (InstantiationException | IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        AxolotlWriteResult axolotlWriteResult;
        if(isFirstBatch()){
            //读取默认配置
            globalCellStyle = getDefaultCellStyle();
            CellMain gsc = new CellMain();
            cellStyleConfigur.globalStyleConfig(gsc);
            setCellStyle(gsc,globalCellStyle);
            String fontName = writeConfig.getFontName();
            if (fontName != null){
                debug(LOGGER, "使用自定义字体：%s",fontName);
                globalCellStyle.setFontName(fontName);
            }

            //读取系统列样式配置
            CellMain csc = new CellMain();
            cellStyleConfigur.commonStyleConfig(csc);
            GlobalCellStyle commonStyle = new GlobalCellStyle();
            try {
                BeanUtils.copyProperties(commonStyle,globalCellStyle);
            } catch (Exception e) {throw new AxolotlWriteException("读取系统列样式配置失败");}
            setCellStyle(csc,commonStyle);
            if(commonStyle.getRowHeight() == null){
                commonStyle.setRowHeight(this.commonRowHeight);
            }
            if(commonStyle.getColumnWidth() == null){
                commonStyle.setColumnWidth(columnWidth);
            }
            this.commonCellStyle = commonStyle;

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

        //读取表头配置
        CellMain hsc = new CellMain();
        cellStyleConfigur.headerStyleConfig(hsc);
        GlobalCellStyle handlerStyle = new GlobalCellStyle();
        try {
            BeanUtils.copyProperties(handlerStyle,globalCellStyle);
        } catch (Exception e) {throw new AxolotlWriteException("读取表头配置失败");}
        setCellStyle(hsc,handlerStyle);
        if(handlerStyle.getRowHeight() == null){
            handlerStyle.setRowHeight(this.headerRowHeight);
        }
        this.handlerCellStyle = handlerStyle;

        //读取标题配置
        CellMain tsc = new CellMain();
        cellStyleConfigur.titleStyleConfig(tsc);
        GlobalCellStyle titleStyle = new GlobalCellStyle();
        try {
            BeanUtils.copyProperties(titleStyle,globalCellStyle);
        } catch (Exception e) {throw new AxolotlWriteException("读取标题配置失败");}
        setCellStyle(tsc,titleStyle);
        if(titleStyle.getRowHeight() == null){
            titleStyle.setRowHeight(this.titleRowHeight);
        }
        this.titleCellStyle = titleStyle;


        // 1.创建标题行
        AxolotlWriteResult writeTitle = createTitleRow(sheet);

        // 2.创建表头单元格样式
        XSSFCellStyle headerDefaultCellStyle = (XSSFCellStyle) createCellStyle(handlerCellStyle.getBaseBorderStyle(), handlerCellStyle.getBaseBorderColor(), handlerCellStyle.getForegroundColor(), handlerCellStyle.getFontName(), handlerCellStyle.getFontSize(), handlerCellStyle.getBold(), handlerCellStyle.getFontColor(), handlerCellStyle.getItalic(), handlerCellStyle.getStrikeout());
        StyleHelper.setCellStyleAlignment(headerDefaultCellStyle, handlerCellStyle.getHorizontalAlignment(), handlerCellStyle.getVerticalAlignment());
        setBorderStyle(headerDefaultCellStyle,handlerCellStyle);
        headerDefaultCellStyle.setWrapText(true);
        headerDefaultCellStyle.setFillPattern(handlerCellStyle.getFillPatternType());

        // 3.渲染表头
        AxolotlWriteResult headerWriteResult = this.defaultRenderHeaders(sheet, headerDefaultCellStyle);

        // 4.创建标题单元格样式
        XSSFCellStyle titleDefaultCellStyle = (XSSFCellStyle) createCellStyle(titleCellStyle.getBaseBorderStyle(), titleCellStyle.getBaseBorderColor(), titleCellStyle.getForegroundColor(), titleCellStyle.getFontName(), titleCellStyle.getFontSize(), titleCellStyle.getBold(), titleCellStyle.getFontColor(), titleCellStyle.getItalic(), titleCellStyle.getStrikeout());
        StyleHelper.setCellStyleAlignment(titleDefaultCellStyle, titleCellStyle.getHorizontalAlignment(), titleCellStyle.getVerticalAlignment());
        setBorderStyle(titleDefaultCellStyle,titleCellStyle);
        titleDefaultCellStyle.setWrapText(true);
        titleDefaultCellStyle.setFillPattern(titleCellStyle.getFillPatternType());

        // 5.合并表头
        if (writeTitle.isWrite()){
            this.mergeTitleRegion(sheet,context.getAlreadyWrittenColumns().get(context.getSwitchSheetIndex()),titleDefaultCellStyle);
        }

        // 5.创建冻结窗格
        sheet.createFreezePane(START_POSITION, context.getAlreadyWriteRow().get(context.getSwitchSheetIndex())+1);

        return headerWriteResult;
    }

    private Map<Integer, Integer> unmappedColumnCount;
    private boolean alreadyNotice = false;
    @Override
    public AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data) {
        //创建系统列单元格样式 用于序号与空值填充
        XSSFCellStyle defaultCellStyle = (XSSFCellStyle) createCellStyle(commonCellStyle.getBaseBorderStyle(), commonCellStyle.getBaseBorderColor(), commonCellStyle.getForegroundColor(), commonCellStyle.getFontName(), commonCellStyle.getFontSize(), commonCellStyle.getBold(), commonCellStyle.getFontColor(), commonCellStyle.getItalic(), commonCellStyle.getStrikeout());
        StyleHelper.setCellStyleAlignment(defaultCellStyle, commonCellStyle.getHorizontalAlignment(), commonCellStyle.getVerticalAlignment());
        setBorderStyle(defaultCellStyle,commonCellStyle);
        defaultCellStyle.setWrapText(true);
        defaultCellStyle.setFillPattern(commonCellStyle.getFillPatternType());
        for (Object datum : data) {
            // 获取对象属性
            HashMap<String, Object> dataMap = new LinkedHashMap<>();
            if (datum instanceof Map map) {
                dataMap.putAll(map);
            }else{
                List<Field> fields = ReflectToolkit.getAllFields(datum.getClass(), true);
                fields.forEach(field -> {
                    field.setAccessible(true);
                    String fieldName = field.getName();
                    try {
                        dataMap.put(fieldName,field.get(datum));
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
            int serialNumber = context.getAndIncrementSerialNumber() - context.getHeaderRowCount().get(switchSheetIndex)+1;
            // 写入序号
            if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER)){
                //写入序号只能设置列宽 行高继承内容
                //设置列宽
                sheet.setColumnWidth(writtenColumn,commonCellStyle.getColumnWidth());
                SXSSFCell cell = dataRow.createCell(writtenColumn);
                cell.setCellValue(serialNumber);
                cell.setCellStyle(defaultCellStyle);
                writtenColumnMap.put(writtenColumn++,1);
            }
            // 写入数据
            Map<String, Integer> columnMapping = context.getHeaderColumnIndexMapping().row(context.getSwitchSheetIndex());
            unmappedColumnCount =  new HashMap<>();
            columnMapping.forEach((key, value) -> unmappedColumnCount.put(value, 1));
            boolean columnMappingEmpty = columnMapping.isEmpty();
            boolean useOrderField = true;
            for (Map.Entry<String, Object> dataEntry : dataMap.entrySet()) {
                String fieldName = dataEntry.getKey();
                SXSSFCell cell;
                if (columnMappingEmpty){
                    cell = dataRow.createCell(writtenColumn);
                }else{
                    useOrderField = false;
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
                FieldInfo fieldInfo = new FieldInfo(fieldName, value, writtenColumn,alreadyWriteRow);

                //获取数据样式
                CellMain dsc = new CellMain();
                cellStyleConfigur.dataStyleConfig(dsc,new FieldInfo(fieldName, value, writtenColumn,alreadyWriteRow));
                GlobalCellStyle dataStyle = new GlobalCellStyle();
                try {
                    BeanUtils.copyProperties(dataStyle,globalCellStyle);
                } catch (Exception e) {throw new AxolotlWriteException("读取数据配置失败");}
                setCellStyle(dsc,dataStyle);
                if(dataStyle.getRowHeight() == null){
                    dataStyle.setRowHeight(this.dataRowHeight);
                }
                List<Header> headers = context.getHeaders().get(switchSheetIndex);
                if(dataStyle.getColumnWidth() == null && (headers == null || headers.isEmpty())){
                    dataStyle.setColumnWidth(columnWidth);
                }

                //创建数据单元格样式
                XSSFCellStyle dataDefaultCellStyle = (XSSFCellStyle) createCellStyle(dataStyle.getBaseBorderStyle(), dataStyle.getBaseBorderColor(), dataStyle.getForegroundColor(), dataStyle.getFontName(), dataStyle.getFontSize(), dataStyle.getBold(), dataStyle.getFontColor(), dataStyle.getItalic(), dataStyle.getStrikeout());
                StyleHelper.setCellStyleAlignment(dataDefaultCellStyle, dataStyle.getHorizontalAlignment(), dataStyle.getVerticalAlignment());
                setBorderStyle(dataDefaultCellStyle,dataStyle);
                dataDefaultCellStyle.setWrapText(true);
                dataDefaultCellStyle.setFillPattern(dataStyle.getFillPatternType());

                //设置行高
                dataRow.setHeight(dataStyle.getRowHeight());
                //设置列宽 内容列宽若没设置默认值，继承表头
                if(dataStyle.getColumnWidth() != null){
                    sheet.setColumnWidth(writtenColumn,dataStyle.getColumnWidth());
                }

                // 对单元格设置样式
                cell.setCellStyle(dataDefaultCellStyle);
                // 渲染数据到单元格
                this.renderColumn(fieldInfo,cell);
                writtenColumnMap.put(writtenColumn++,1);
            }
            // 将未使用的的单元格赋予空值
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
                    cell.setCellStyle(defaultCellStyle);
                    //空值填充不设置行高列宽
                }
            }
        }

        return new AxolotlWriteResult(true, "渲染数据完成");
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
            row.setHeight(commonCellStyle.getRowHeight());
            DataInverter<?> dataInverter = writeConfig.getDataInverter();
            HashSet<Integer> calculateColumnIndexes = writeConfig.getCalculateColumnIndexes();
            for (int i = 0; i < alreadyWrittenColumns; i++) {
                SXSSFCell cell = row.createCell(i);
                String cellValue = "-";
                if (i == 0 && writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER)){
                    cellValue = "合计";
                    sheet.setColumnWidth(i, commonCellStyle.getColumnWidth());
                }
                if (endingTotalMapping.containsKey(i)){
                    if (calculateColumnIndexes.contains(i) || (calculateColumnIndexes.size() == 1 && calculateColumnIndexes.contains(-1))){
                        BigDecimal bigDecimal = endingTotalMapping.get(i);
                        Object convert = dataInverter.convert(bigDecimal);
                        cellValue = convert.toString();
                    }
                }
                cell.setCellValue(cellValue);
                //合计部分样式继承上一行
                CellStyle cellStyle = sheet.getRow(sheet.getLastRowNum() - 1).getCell(i).getCellStyle();
                SXSSFWorkbook workbook = sheet.getWorkbook();
                CellStyle totalCellStyle = workbook.createCellStyle();
                Font font = StyleHelper.createWorkBookFont(
                        workbook, commonCellStyle.getFontName(), commonCellStyle.getBold(), commonCellStyle.getFontSize(), commonCellStyle.getFontColor(), commonCellStyle.getItalic(), commonCellStyle.getStrikeout()
                );
                StyleHelper.setCellStyleAlignment(totalCellStyle,commonCellStyle.getHorizontalAlignment(),commonCellStyle.getVerticalAlignment());
                totalCellStyle.setFillForegroundColor(cellStyle.getFillForegroundColorColor());
                totalCellStyle.setFillPattern(cellStyle.getFillPattern());
                totalCellStyle.setFont(font);

                totalCellStyle.setBorderBottom(cellStyle.getBorderBottom());
                totalCellStyle.setBorderLeft(cellStyle.getBorderLeft());
                totalCellStyle.setBorderRight(cellStyle.getBorderRight());
                totalCellStyle.setBorderTop(cellStyle.getBorderTop());
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
                sheet.trackAllColumnsForAutoSizing();
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


    /**
     * 创建预制样式
     * @return
     */
    private GlobalCellStyle getDefaultCellStyle(){
        GlobalCellStyle defaultStyle = new GlobalCellStyle();

        defaultStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        defaultStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        defaultStyle.setForegroundColor(new AxolotlColor(255,255,255));
        defaultStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        //边框
        defaultStyle.setBaseBorderStyle(BorderStyle.NONE);
        defaultStyle.setBaseBorderColor(IndexedColors.BLACK);
        //字体
        defaultStyle.setFontName(StyleHelper.STANDARD_FONT_NAME);
        defaultStyle.setFontColor(IndexedColors.BLACK);
        defaultStyle.setBold(false);
        defaultStyle.setFontSize(StyleHelper.STANDARD_TEXT_FONT_SIZE);
        defaultStyle.setItalic(false);
        defaultStyle.setStrikeout(false);

        return defaultStyle;
    }

    /**
     * 导入配置
     * @param cellMain 用户配置
     * @param defaultCellStyle 主题配置
     */
    private void setCellStyle(CellMain cellMain, GlobalCellStyle defaultCellStyle){
        if(cellMain.getRowHeight() != null){
            defaultCellStyle.setRowHeight(cellMain.getRowHeight());
        }
        if(cellMain.getColumnWidth() != null){
            defaultCellStyle.setColumnWidth(cellMain.getColumnWidth());
        }
        if(cellMain.getHorizontalAlignment() != null){
            defaultCellStyle.setHorizontalAlignment(cellMain.getHorizontalAlignment());
        }
        if(cellMain.getVerticalAlignment() != null){
            defaultCellStyle.setVerticalAlignment(cellMain.getVerticalAlignment());
        }
        if(cellMain.getForegroundColor() != null){
            defaultCellStyle.setForegroundColor(cellMain.getForegroundColor());
        }
        if(cellMain.getFillPatternType() != null){
            defaultCellStyle.setFillPatternType(cellMain.getFillPatternType());
        }
        CellBorder border = cellMain.getBorder();
        if(border != null){
            if(border.getBaseBorderStyle() != null){
                defaultCellStyle.setBaseBorderStyle(border.getBaseBorderStyle());
            }
            if(border.getBaseBorderColor() != null){
                defaultCellStyle.setBaseBorderColor(border.getBaseBorderColor());
            }
            if(border.getBorderTopStyle() != null){
                defaultCellStyle.setBorderTopStyle(border.getBorderTopStyle());
            }
            if(border.getTopBorderColor() != null){
                defaultCellStyle.setTopBorderColor(border.getTopBorderColor());
            }
            if(border.getBorderBottomStyle() != null){
                defaultCellStyle.setBorderBottomStyle(border.getBorderBottomStyle());
            }
            if(border.getBottomBorderColor() != null){
                defaultCellStyle.setBottomBorderColor(border.getBottomBorderColor());
            }
            if(border.getBorderLeftStyle() != null){
                defaultCellStyle.setBorderLeftStyle(border.getBorderLeftStyle());
            }
            if(border.getLeftBorderColor() != null){
                defaultCellStyle.setLeftBorderColor(border.getLeftBorderColor());
            }
            if(border.getBorderRightStyle() != null){
                defaultCellStyle.setBorderRightStyle(border.getBorderRightStyle());
            }
            if(border.getRightBorderColor() != null){
                defaultCellStyle.setRightBorderColor(border.getRightBorderColor());
            }
        }
        CellFont font = cellMain.getFont();
        if(font != null){
            if(font.getBold() != null){
                defaultCellStyle.setBold(font.getBold());
            }
            if(font.getFontName() != null){
                defaultCellStyle.setFontName(font.getFontName());
            }
            if(font.getFontSize() != null){
                defaultCellStyle.setFontSize(font.getFontSize());
            }
            if(font.getFontColor() != null){
                defaultCellStyle.setFontColor(font.getFontColor());
            }
            if(font.getItalic() != null){
                defaultCellStyle.setItalic(font.getItalic());
            }
            if(font.getStrikeout() != null){
                defaultCellStyle.setStrikeout(font.getStrikeout());
            }
        }
    }

    /**
     * 填充空白单元格
     * @param sheet 工作表
     */
    public void fillWhiteCell(Sheet sheet){
       // Font font = createFont(globalCellStyle.getFontName(), globalCellStyle.getFontSize(), globalCellStyle.isBold(), globalCellStyle.getFontColor());
        CellStyle defaultStyle = createCellStyle(globalCellStyle.getBaseBorderStyle(), globalCellStyle.getBaseBorderColor(), globalCellStyle.getForegroundColor(), globalCellStyle.getFontName(), globalCellStyle.getFontSize(), globalCellStyle.getBold(), globalCellStyle.getFontColor(), globalCellStyle.getItalic(), globalCellStyle.getStrikeout());
        //设置单元格对齐方式
        StyleHelper.setCellStyleAlignment(defaultStyle, globalCellStyle.getHorizontalAlignment(), globalCellStyle.getVerticalAlignment());
        //设置单元格边框样式
        setBorderStyle(defaultStyle,globalCellStyle);
        //填充样式
        defaultStyle.setFillPattern(globalCellStyle.getFillPatternType());
        // 将默认样式应用到所有单元格
        for (int i = 0; i < 26; i++) {
            sheet.setDefaultColumnStyle(i, defaultStyle);
            sheet.setDefaultColumnWidth(globalCellStyle.getColumnWidth() == null ? columnWidth : globalCellStyle.getColumnWidth());
        }
        sheet.setDefaultRowHeight(globalCellStyle.getRowHeight() == null ? dataRowHeight : globalCellStyle.getRowHeight());
    }

    /**
     * 根据全局配置设置单元格边框样式
     * @param cellStyle 单元格样式
     * @param baseCellStyle 全局配置
     */
    private void setBorderStyle(CellStyle cellStyle, GlobalCellStyle baseCellStyle){
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
            titleRow.setHeight(titleCellStyle.getRowHeight());
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
                row.setHeight(handlerCellStyle.getRowHeight());

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
                    headerRecursiveInfo.setRowHeight(handlerCellStyle.getRowHeight());
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
                            if(handlerCellStyle.getColumnWidth() != null){
                                columnWidth = handlerCellStyle.getColumnWidth();
                            }else{
                                columnWidth = StyleHelper.getPresetCellLength(title);
                            }
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
                Font font = createFont(handlerCellStyle.getFontName(), handlerCellStyle.getFontSize(), handlerCellStyle.getBold(), handlerCellStyle.getFontColor());
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

    /**
     * 创建样式、从缓存中获取样式
     * @param borderStyle 边框样式
     * @param borderColor 边框颜色
     * @param cellColor 单元格背景色
     * @param fontName 字体名称
     * @param fontSize 字体大小
     * @param isBold 是否加粗
     * @param fontColor 字体颜色
     * @param italic 是否斜体
     * @param strikeout 是否删除线
     * @return 样式
     */
    public CellStyle createCellStyle(BorderStyle borderStyle, IndexedColors borderColor, AxolotlColor cellColor,
                                     String fontName,Short fontSize,Boolean isBold,Object fontColor, Boolean italic,Boolean strikeout){
        GlobalCellStyle key = new GlobalCellStyle();
        key.setBaseBorderStyle(borderStyle);
        key.setBaseBorderColor(borderColor);
        key.setForegroundColor(cellColor);
        key.setFontName(fontName);
        key.setFontSize(fontSize);
        key.setBold(isBold);
        key.setFontColor((IndexedColors) fontColor);
        key.setItalic(italic);
        key.setStrikeout(strikeout);

        if(cellStyleCache.containsKey(key)){
            return cellStyleCache.get(key);
        }else{
            CellStyle style = createStyle(borderStyle, borderColor, cellColor,
                    fontName, fontSize, isBold, fontColor, italic, strikeout);
            cellStyleCache.put(key,style);
            return style;
        }
    }
}

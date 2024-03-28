package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.BaseCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.CellStyleConfigur;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.checkerframework.checker.units.qual.A;
import org.slf4j.Logger;

import java.lang.reflect.Field;
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


    private BaseCellStyle getDefaultCellStyle(){
        BaseCellStyle baseCellStyle = new BaseCellStyle();
        baseCellStyle.setTitleRowHeight(StyleHelper.STANDARD_TITLE_ROW_HEIGHT);
        baseCellStyle.setHeaderRowHeight(StyleHelper.STANDARD_HEADER_ROW_HEIGHT);
        baseCellStyle.setDataRowHeight(StyleHelper.STANDARD_HEADER_ROW_HEIGHT);

        baseCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        baseCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        baseCellStyle.setForegroundColor(new AxolotlColor(255,255,255));

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

}

package cn.xvoid.axolotl.excel.writer.themes.standard;

import cn.xvoid.axolotl.excel.writer.components.configuration.AxolotlColor;
import cn.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.axolotl.toolkit.LoggerHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;

import java.util.List;
import java.util.Map;

import static cn.xvoid.axolotl.excel.writer.style.StyleHelper.START_POSITION;
import static cn.xvoid.axolotl.toolkit.LoggerHelper.debug;

public class AxolotlSimpleBlackTheme extends AbstractStyleRender implements ExcelStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlSimpleBlackTheme.class);

    private static final AxolotlColor THEME_COLOR_XSSF = new AxolotlColor(255,255,255);

    private static final String FONT_NAME = "宋体";

    private Font MAIN_TEXT_FONT;

    public AxolotlSimpleBlackTheme() {
        super(LOGGER);
    }
    @Override
    public AxolotlWriteResult init(Sheet sheet) {
        if (isFirstBatch()){
            this.checkedAndUseCustomTheme(FONT_NAME,THEME_COLOR_XSSF);
            MAIN_TEXT_FONT = StyleHelper.createWorkBookFont(context.getWorkbook(), FONT_NAME, false, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.BLACK);
        }
        return super.init(sheet);
    }

    @Override
    public AxolotlWriteResult renderHeader(Sheet sheet) {
        // 1.渲染标题
        int switchSheetIndex = context.getSwitchSheetIndex();
        AxolotlWriteResult isCreateTitleRow = this.createTitleRow(sheet);

        // 2.渲染表头
        Font headerFont = StyleHelper.createWorkBookFont(context.getWorkbook(), FONT_NAME, true, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.BLACK);
        CellStyle headerDefaultCellStyle = StyleHelper.createStandardCellStyle(context.getWorkbook(), BorderStyle.THIN, IndexedColors.BLACK, THEME_COLOR_XSSF,headerFont);
        AxolotlWriteResult headerWriteResult = this.defaultRenderHeaders(sheet, headerDefaultCellStyle);

        // 3.合并标题列单元格并赋予样式
        if (isCreateTitleRow.isWrite()){
            Font titleFont = StyleHelper.createWorkBookFont(context.getWorkbook(), FONT_NAME, true, StyleHelper.STANDARD_TITLE_FONT_SIZE, IndexedColors.BLACK);
            CellStyle titleRowStyle = StyleHelper.createStandardCellStyle(context.getWorkbook(), BorderStyle.THIN, IndexedColors.BLACK, THEME_COLOR_XSSF,titleFont);
            this.mergeTitleRegion(sheet,context.getAlreadyWrittenColumns().get(switchSheetIndex),titleRowStyle);
        }

        // 4.创建冻结窗格
        sheet.createFreezePane(START_POSITION, context.getAlreadyWriteRow().get(switchSheetIndex)+1);

        return headerWriteResult;
    }

    @Override
    public AxolotlWriteResult renderData(Sheet sheet, List<?> data) {
        Workbook workbook = context.getWorkbook();
        BorderStyle borderStyle = BorderStyle.THIN;
        IndexedColors borderColor = IndexedColors.BLACK;
        // 交叉样式
        CellStyle dataStyle = StyleHelper.createStandardCellStyle(workbook, borderStyle, borderColor, THEME_COLOR_XSSF,MAIN_TEXT_FONT);
        StyleHelper.setCellAsPlainText(dataStyle);
        Map<String, Integer> columnMapping = context.getHeaderColumnIndexMapping().row(context.getSwitchSheetIndex());
        if (!columnMapping.isEmpty()){
            LoggerHelper.debug(LOGGER,"已有字段映射表,将按照字段映射渲染数据[%s]",columnMapping);
        }
        for (Object datum : data) {
            this.defaultRenderNextData(sheet, datum, dataStyle);
        }
        return new AxolotlWriteResult(true, "渲染数据完成");
    }

    @Override
    public AxolotlWriteResult finish(Sheet sheet) {
        return super.finish(sheet);
    }

}

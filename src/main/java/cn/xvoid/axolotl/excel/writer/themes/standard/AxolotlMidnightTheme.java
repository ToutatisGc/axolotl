package cn.xvoid.axolotl.excel.writer.themes.standard;

import cn.xvoid.axolotl.excel.writer.components.configuration.AxolotlColor;
import cn.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.axolotl.toolkit.LoggerHelper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.slf4j.Logger;

import java.util.List;
import java.util.Map;

import static cn.xvoid.axolotl.toolkit.LoggerHelper.debug;

public class AxolotlMidnightTheme extends AbstractStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlMidnightTheme.class);

    private static final String FONT_NAME = "Calibri";
    public AxolotlMidnightTheme() {
        super(LOGGER);
    }

    @Override
    public AxolotlWriteResult init(Sheet sheet) {
        if (isFirstBatch()){
            this.checkedAndUseCustomTheme(FONT_NAME,null);
        }
        return super.init(sheet);
    }

    @Override
    public AxolotlWriteResult renderHeader(Sheet sheet) {
        // 1.创建标题行
        AxolotlWriteResult writeTitle = createTitleRow(sheet);

        // 2.创建表头单元格样式
        Font headerFont = createWhiteMainTextFont();
        headerFont.setBold(true);
        CellStyle headerDefaultCellStyle = this.createStyle(BorderStyle.THIN, IndexedColors.BLACK, AxolotlColor.create(34,44,71), headerFont);
        headerDefaultCellStyle.setBorderTop(BorderStyle.NONE);
        headerDefaultCellStyle.setBorderBottom(BorderStyle.NONE);
        headerDefaultCellStyle.setWrapText(true);
        // 3.渲染表头
        AxolotlWriteResult headerWriteResult = this.defaultRenderHeaders(sheet, headerDefaultCellStyle);

        // 4.合并表头
        if (writeTitle.isWrite()){
            Font titleFont = createWhiteMainTextFont();
            titleFont.setFontHeightInPoints(StyleHelper.STANDARD_TITLE_FONT_SIZE);
            CellStyle titleStyle = createStyle(BorderStyle.NONE, IndexedColors.BLACK, AxolotlColor.create(53, 80, 125), titleFont);
            this.mergeTitleRegion(sheet,context.getAlreadyWrittenColumns().get(context.getSwitchSheetIndex()),titleStyle);
        }

        return headerWriteResult;
    }

    @Override
    public AxolotlWriteResult renderData(Sheet sheet, List<?> data) {
        CellStyle dataStyle = this.createWhiteMainTextCellStyle(BorderStyle.NONE, IndexedColors.BLACK, AxolotlColor.create(39,56,86));
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setBorderBottom(BorderStyle.THIN);
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
    public void fillWhiteCell(Sheet sheet, String fontName) {
        CellStyle defaultStyle = createWhiteMainTextCellStyle(BorderStyle.NONE, IndexedColors.WHITE, AxolotlColor.create(52, 64, 90));
        // 将默认样式应用到所有单元格
        for (int i = 0; i < 26; i++) {
            sheet.setDefaultColumnStyle(i, defaultStyle);
            sheet.setDefaultColumnWidth(12);
        }
        sheet.setDefaultRowHeight(StyleHelper.STANDARD_ROW_HEIGHT);
    }
}

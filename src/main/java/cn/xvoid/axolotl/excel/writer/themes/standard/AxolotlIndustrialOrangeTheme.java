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

import org.slf4j.Logger;

import java.util.List;
import java.util.Map;

import static cn.xvoid.axolotl.toolkit.LoggerHelper.debug;

public class AxolotlIndustrialOrangeTheme extends AbstractStyleRender implements ExcelStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlIndustrialOrangeTheme.class);

    private static final AxolotlColor THEME_COLOR = AxolotlColor.create(247,202,142);

    private static final String FONT_NAME = "Calibri";

    public AxolotlIndustrialOrangeTheme() {
        super(LOGGER);
    }

    @Override
    public AxolotlWriteResult init(Sheet sheet) {
        if (isFirstBatch()){
            this.checkedAndUseCustomTheme(FONT_NAME,THEME_COLOR);
        }
        return super.init(sheet);
    }

    @Override
    public AxolotlWriteResult renderHeader(Sheet sheet) {
        // 1.创建标题行
        AxolotlWriteResult writeTitle = createTitleRow(sheet);

        // 2.创建表头单元格样式
        Font font = this.createFont(getGlobalFontName(), StyleHelper.STANDARD_TEXT_FONT_SIZE, true, IndexedColors.BLACK);
        CellStyle headerDefaultCellStyle = this.createStyle(BorderStyle.MEDIUM, IndexedColors.BLACK, THEME_COLOR, font);
        headerDefaultCellStyle.setWrapText(true);
        // 3.渲染表头
        AxolotlWriteResult headerWriteResult = this.defaultRenderHeaders(sheet, headerDefaultCellStyle);

        // 4.合并表头
        if (writeTitle.isWrite()){
            this.mergeTitleRegion(sheet,context.getAlreadyWrittenColumns().get(context.getSwitchSheetIndex()),headerDefaultCellStyle);
        }

        return headerWriteResult;
    }

    @Override
    public AxolotlWriteResult renderData(Sheet sheet, List<?> data) {
        CellStyle dataStyle = this.createBlackMainTextCellStyle(BorderStyle.THIN, IndexedColors.BLACK, StyleHelper.WHITE_COLOR);
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

package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;

import java.util.List;
import java.util.Map;

import static cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper.START_POSITION;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.debug;

public class SimpleBlackTheme extends AbstractStyleRender implements ExcelStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(SimpleBlackTheme.class);

    private static final AxolotlColor THEME_COLOR_XSSF = new AxolotlColor(255,255,255);

    private static final String FONT_NAME = "宋体";

    private Font MAIN_TEXT_FONT;

    public SimpleBlackTheme() {
        super(LOGGER);
    }
    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        if (isFirstBatch()){
            MAIN_TEXT_FONT = StyleHelper.createWorkBookFont(context.getWorkbook(), FONT_NAME, false, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.BLACK);
        }
        return super.init(sheet);
    }

    @Override
    public AxolotlWriteResult renderHeader(SXSSFSheet sheet) {
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
    public AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data) {
        SXSSFWorkbook workbook = context.getWorkbook();
        BorderStyle borderStyle = BorderStyle.THIN;
        IndexedColors borderColor = IndexedColors.BLACK;
        // 交叉样式
        CellStyle dataStyle = StyleHelper.createStandardCellStyle(workbook, borderStyle, borderColor, THEME_COLOR_XSSF,MAIN_TEXT_FONT);
        StyleHelper.setCellAsPlainText(dataStyle);
        Map<String, Integer> columnMapping = context.getHeaderColumnIndexMapping().row(context.getSwitchSheetIndex());
        if (!columnMapping.isEmpty()){
            debug(LOGGER,"已有字段映射表,将按照字段映射渲染数据[%s]",columnMapping);
        }
        for (int i = 0, dataSize = data.size(); i < dataSize; i++) {
            this.defaultRenderNextData(sheet, data.get(i), dataStyle);
        }
        return new AxolotlWriteResult(true, "渲染数据完成");
    }

    @Override
    public AxolotlWriteResult finish(SXSSFSheet sheet) {
        return super.finish(sheet);
    }

}

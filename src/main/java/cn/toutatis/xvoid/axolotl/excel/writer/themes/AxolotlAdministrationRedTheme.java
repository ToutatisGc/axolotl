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
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.IndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;

import java.util.List;

public class AxolotlAdministrationRedTheme extends AbstractStyleRender implements ExcelStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlAdministrationRedTheme.class);

    private static final AxolotlColor THEME_COLOR_XSSF = new AxolotlColor(255,255,255);

    private static final String FONT_NAME = "黑体";

    private Font MAIN_TEXT_FONT;

    public AxolotlAdministrationRedTheme() {
        super(LOGGER,FONT_NAME);
    }

    @Override
    public AxolotlWriteResult renderHeader(SXSSFSheet sheet) {
        AxolotlWriteResult writeTitle = createTitleRow(sheet);
        // 创建调色板
        StylesTable stylesSource = sheet.getWorkbook().getXSSFWorkbook().getStylesSource();
        XSSFFont xssfFont = new XSSFFont();
        xssfFont.setColor(new AxolotlColor(255,0,255).toXSSFColor());
        SXSSFWorkbook workbook = context.getWorkbook();
        xssfFont.registerTo(stylesSource);
        CellStyle cellStyle = workbook.createCellStyle();
        workbook.createFont();
        cellStyle.setFont(xssfFont);
        // 2.渲染表头
        Font headerFont = StyleHelper.createWorkBookFont(context.getWorkbook(), FONT_NAME, true, StyleHelper.STANDARD_TEXT_FONT_SIZE, IndexedColors.RED);
        CellStyle headerDefaultCellStyle = StyleHelper.createStandardCellStyle(context.getWorkbook(), BorderStyle.THIN, IndexedColors.WHITE, THEME_COLOR_XSSF,headerFont);
        AxolotlWriteResult headerWriteResult = this.defaultRenderHeaders(sheet, cellStyle);

        if (writeTitle.isWrite()){
            SXSSFRow titleRow = sheet.getRow(StyleHelper.START_POSITION);
        }

        return super.init(sheet);
    }

    @Override
    public AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data) {
        return null;
    }
}


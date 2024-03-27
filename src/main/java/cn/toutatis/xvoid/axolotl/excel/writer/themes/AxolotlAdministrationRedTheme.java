package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;

import java.util.List;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

public class AxolotlAdministrationRedTheme extends AbstractStyleRender implements ExcelStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlAdministrationRedTheme.class);

    private static final AxolotlColor THEME_COLOR_XSSF = new AxolotlColor(191,56,55);

    private static final String THEME_FONT_NAME = "黑体";

    private Font MAIN_TEXT_FONT;

    public AxolotlAdministrationRedTheme() {
        super(LOGGER);
    }

    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        if (isFirstBatch()){
            // 固定结构
            super.checkedAndUseCustomFontName(THEME_FONT_NAME);
            // 创建调色板
            XSSFColor themeColorXssfxssfColor = THEME_COLOR_XSSF.toXSSFColor();
            StylesTable stylesSource = context.getWorkbook().getXSSFWorkbook().getStylesSource();
            XSSFCellStyle xssfCellStyle = new XSSFCellStyle(stylesSource);
            xssfCellStyle.setBottomBorderColor(themeColorXssfxssfColor);
            stylesSource.putStyle(xssfCellStyle);
            XSSFFont mainFont = new XSSFFont();
            mainFont.setBold(true);
            mainFont.setColor(themeColorXssfxssfColor);
            mainFont.setFontName(getFontName());
            mainFont.setFontHeightInPoints(StyleHelper.STANDARD_TEXT_FONT_SIZE);
            mainFont.registerTo(stylesSource);
            MAIN_TEXT_FONT = mainFont;
        }
        return super.init(sheet);
    }

    @Override
    public AxolotlWriteResult renderHeader(SXSSFSheet sheet) {
        // 1.创建标题行
        AxolotlWriteResult writeTitle = createTitleRow(sheet);

        // 2.创建表头单元格样式
        AxolotlColor whiteColor = new AxolotlColor(255, 255, 255);
        XSSFCellStyle headerDefaultCellStyle = (XSSFCellStyle) StyleHelper.
                createStandardCellStyle(context.getWorkbook(), BorderStyle.NONE, IndexedColors.WHITE,whiteColor,MAIN_TEXT_FONT);
        headerDefaultCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        headerDefaultCellStyle.setBottomBorderColor(THEME_COLOR_XSSF.toXSSFColor());
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
    public AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data) {
        return null;
    }

    @Override
    public AxolotlWriteResult finish(SXSSFSheet sheet) {
//        return super.finish(sheet);
        return null;
    }
}


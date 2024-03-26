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
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.slf4j.Logger;

import java.util.List;

public class AxolotlAdministrationRedTheme extends AbstractStyleRender implements ExcelStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlAdministrationRedTheme.class);

    private static final AxolotlColor THEME_COLOR_XSSF = new AxolotlColor(191,56,55);

    private static final String FONT_NAME = "黑体";

    private Font MAIN_TEXT_FONT;

    public AxolotlAdministrationRedTheme() {
        super(LOGGER,FONT_NAME);
    }

    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        if (isFirstBatch()){
            // 创建调色板
            StylesTable stylesSource = context.getWorkbook().getXSSFWorkbook().getStylesSource();
            XSSFCellStyle xssfCellStyle = new XSSFCellStyle(stylesSource);
            xssfCellStyle.setBottomBorderColor(THEME_COLOR_XSSF.toXSSFColor());
            stylesSource.putStyle(xssfCellStyle);
            XSSFFont mainFont = new XSSFFont();
            mainFont.setBold(true);
            mainFont.setColor(THEME_COLOR_XSSF.toXSSFColor());
            mainFont.setFontName(FONT_NAME);
            mainFont.registerTo(stylesSource);
            MAIN_TEXT_FONT = mainFont;
        }
        return super.init(sheet);
    }

    @Override
    public AxolotlWriteResult renderHeader(SXSSFSheet sheet) {
        AxolotlWriteResult writeTitle = createTitleRow(sheet);

        // 2.渲染表头
        AxolotlColor whiteColor = new AxolotlColor(255, 255, 255);
        XSSFCellStyle headerDefaultCellStyle = (XSSFCellStyle) StyleHelper.createStandardCellStyle(context.getWorkbook(), BorderStyle.THIN, IndexedColors.WHITE,whiteColor,MAIN_TEXT_FONT);
        headerDefaultCellStyle.setBottomBorderColor(IndexedColors.RED.index);
        headerDefaultCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        headerDefaultCellStyle.setBottomBorderColor(THEME_COLOR_XSSF.toXSSFColor());
        headerDefaultCellStyle.setWrapText(true);
        AxolotlWriteResult headerWriteResult = this.defaultRenderHeaders(sheet, headerDefaultCellStyle);

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


package cn.toutatis.xvoid.axolotl.excel.writer.themes.standard;

import cn.toutatis.xvoid.axolotl.excel.writer.components.configuration.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.xvoid.toolkit.log.LoggerToolkit;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.slf4j.Logger;

import java.util.List;
import java.util.Map;

import static cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper.START_POSITION;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

public class AxolotlAdministrationRedTheme extends AbstractStyleRender implements ExcelStyleRender {

    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlAdministrationRedTheme.class);
    /**
     * 主题色
     */
    private static final AxolotlColor THEME_COLOR = AxolotlColor.create(255, 255, 255);
    /**
     * 主题字体颜色
     */
    private static final AxolotlColor THEME_FONT_COLOR = AxolotlColor.create(191,56,55);
    /**
     * 主题字体名称
     */
    private static final String THEME_FONT_NAME = "仿宋";

    private Font MAIN_TEXT_FONT;

    public AxolotlAdministrationRedTheme() {
        super(LOGGER);
    }

    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        AxolotlWriteResult init = super.init(sheet);
        // 固定结构
        if (isFirstBatch()){
            this.checkedAndUseCustomTheme(THEME_FONT_NAME,THEME_COLOR);
            MAIN_TEXT_FONT = this.createFont(getGlobalFontName(), StyleHelper.STANDARD_TEXT_FONT_SIZE, true, THEME_FONT_COLOR);
        }
        return init;
    }

    @Override
    public AxolotlWriteResult renderHeader(SXSSFSheet sheet) {
        // 1.创建标题行
        AxolotlWriteResult writeTitle = createTitleRow(sheet);

        // 2.创建表头单元格样式
        XSSFCellStyle headerDefaultCellStyle = (XSSFCellStyle) this.createStyle(BorderStyle.NONE, IndexedColors.WHITE, THEME_COLOR, MAIN_TEXT_FONT);
        headerDefaultCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        headerDefaultCellStyle.setBottomBorderColor(THEME_FONT_COLOR.toXSSFColor());
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
        CellStyle dataStyle = this.createBlackMainTextCellStyle(BorderStyle.NONE, IndexedColors.WHITE, THEME_COLOR);
        StyleHelper.setCellAsPlainText(dataStyle);
        Map<String, Integer> columnMapping = context.getHeaderColumnIndexMapping().row(context.getSwitchSheetIndex());
        if (!columnMapping.isEmpty()){
            debug(LOGGER,"已有字段映射表,将按照字段映射渲染数据[%s]",columnMapping);
        }
        for (Object datum : data) {
            this.defaultRenderNextData(sheet, datum, dataStyle);
        }
        return new AxolotlWriteResult(true, "渲染数据完成");
    }

    @Override
    public AxolotlWriteResult finish(SXSSFSheet sheet) {
        return super.finish(sheet);
    }
}


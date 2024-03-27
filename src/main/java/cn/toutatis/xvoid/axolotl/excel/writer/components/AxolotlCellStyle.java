package cn.toutatis.xvoid.axolotl.excel.writer.components;

import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

@Data
public class AxolotlCellStyle {

    /**
     * 前景色
     */
    private AxolotlColor foregroundColor = new AxolotlColor(0,0,0);

    /**
     * 填充模式
     */
    private FillPatternType fillPatternType = FillPatternType.SOLID_FOREGROUND;

    /**
     * 边框颜色
     */
    private IndexedColors borderColor = IndexedColors.BLACK;

    /**
     * 边框样式
     */
    private BorderStyle borderStyle = BorderStyle.THIN;

    /**
     * 字体
     */
    private String fontName = StyleHelper.STANDARD_FONT_NAME;

    /**
     * 字体大小
     */
    private short fontSize = StyleHelper.STANDARD_TEXT_FONT_SIZE;

    /**
     * 字体颜色
     */
    private IndexedColors fontColor = IndexedColors.BLACK;

    /**
     * 字体是否加粗
     */
    private boolean fontBold = false;

}

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
     * 上边框样式
     */
    private BorderStyle borderTopStyle = BorderStyle.THIN;

    /**
     * 上边框颜色
     */
    private IndexedColors topBorderColor = IndexedColors.BLACK;

    /**
     * 下边框样式
     */
    private BorderStyle borderBottomStyle = BorderStyle.THIN;

    /**
     * 下边框颜色
     */
    private IndexedColors bottomBorderColor = IndexedColors.BLACK;

    /**
     * 左边框样式
     */
    private BorderStyle borderLeftStyle = BorderStyle.THIN;

    /**
     * 左边框颜色
     */
    private IndexedColors leftBorderColor = IndexedColors.BLACK;

    /**
     * 右边框样式
     */
    private BorderStyle borderRightStyle = BorderStyle.THIN;

    /**
     * 右边框颜色
     */
    private IndexedColors rightBorderColor = IndexedColors.BLACK;

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

    /**
     *设置文字为斜体
     */
    private boolean italic = false;

    /**
     * 使用水平删除线
     */
    private boolean strikeout = false;

}

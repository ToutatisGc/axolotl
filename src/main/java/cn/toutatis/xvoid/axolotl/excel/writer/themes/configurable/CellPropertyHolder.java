package cn.toutatis.xvoid.axolotl.excel.writer.themes.configurable;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import lombok.Data;
import org.apache.poi.ss.usermodel.*;

import java.util.Objects;

/**
 * 单元格属性容器
 * @author 张智凯
 * @version 1.0
 *  2024/3/28 9:32
 */
@Data
public class CellPropertyHolder {

    /**
     * 行高
     */
    private Short rowHeight;

    /**
     * 列宽
     */
    private Short columnWidth;

    /**
     * 单元格水平对齐方式
     */
    private HorizontalAlignment horizontalAlignment;

    /**
     * 单元格对齐方式
     */
    private VerticalAlignment verticalAlignment;

    /**
     * 背景颜色
     */
    private AxolotlColor foregroundColor;

    /**
     * 填充模式
     */
    private FillPatternType fillPatternType;

    /**
     * 边框默认样式
     */
    private BorderStyle baseBorderStyle;

    /**
     * 边框默认颜色
     */
    private IndexedColors baseBorderColor;

    /**
     * 上边框样式
     */
    private BorderStyle borderTopStyle;

    /**
     * 上边框颜色
     */
    private IndexedColors topBorderColor;

    /**
     * 下边框样式
     */
    private BorderStyle borderBottomStyle;

    /**
     * 下边框颜色
     */
    private IndexedColors bottomBorderColor;

    /**
     * 左边框样式
     */
    private BorderStyle borderLeftStyle;

    /**
     * 左边框颜色
     */
    private IndexedColors leftBorderColor;

    /**
     * 右边框样式
     */
    private BorderStyle borderRightStyle;

    /**
     * 右边框颜色
     */
    private IndexedColors rightBorderColor;

    /**
     * 字体名称
     */
    private String fontName;

    /**
     * 是否加粗
     */
    private Boolean bold;

    /**
     * 字体大小
     */
    private Short fontSize;

    /**
     *字体颜色
     */
    private IndexedColors fontColor;

    /**
     *设置文字为斜体
     */
    private Boolean italic;

    /**
     * 使用水平删除线
     */
    private Boolean strikeout;

    /**
     * 构建默认样式
     * @return 默认样式
     */
    public static CellPropertyHolder buildDefault(){
        CellPropertyHolder defaultStyle = new CellPropertyHolder();

        defaultStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        defaultStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        defaultStyle.setForegroundColor(StyleHelper.WHITE_COLOR);
        defaultStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        //边框
        defaultStyle.setBaseBorderStyle(BorderStyle.NONE);
        defaultStyle.setBaseBorderColor(IndexedColors.BLACK);
        //字体
        defaultStyle.setFontName(StyleHelper.STANDARD_FONT_NAME);
        defaultStyle.setFontColor(IndexedColors.BLACK);
        defaultStyle.setBold(false);
        defaultStyle.setFontSize(StyleHelper.STANDARD_TEXT_FONT_SIZE);
        defaultStyle.setItalic(false);
        defaultStyle.setStrikeout(false);

        return defaultStyle;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        CellPropertyHolder style = (CellPropertyHolder) o;
        return Objects.equals(rowHeight, style.rowHeight) && Objects.equals(columnWidth, style.columnWidth) && horizontalAlignment == style.horizontalAlignment && verticalAlignment == style.verticalAlignment && Objects.equals(foregroundColor, style.foregroundColor) && fillPatternType == style.fillPatternType && baseBorderStyle == style.baseBorderStyle && baseBorderColor == style.baseBorderColor && borderTopStyle == style.borderTopStyle && topBorderColor == style.topBorderColor && borderBottomStyle == style.borderBottomStyle && bottomBorderColor == style.bottomBorderColor && borderLeftStyle == style.borderLeftStyle && leftBorderColor == style.leftBorderColor && borderRightStyle == style.borderRightStyle && rightBorderColor == style.rightBorderColor && Objects.equals(fontName, style.fontName) && Objects.equals(bold, style.bold) && Objects.equals(fontSize, style.fontSize) && fontColor == style.fontColor && Objects.equals(italic, style.italic) && Objects.equals(strikeout, style.strikeout);
    }

    @Override
    public int hashCode() {
        return Objects.hash(rowHeight, columnWidth, horizontalAlignment, verticalAlignment, foregroundColor, fillPatternType, baseBorderStyle, baseBorderColor, borderTopStyle, topBorderColor, borderBottomStyle, bottomBorderColor, borderLeftStyle, leftBorderColor, borderRightStyle, rightBorderColor, fontName, bold, fontSize, fontColor, italic, strikeout);
    }

}

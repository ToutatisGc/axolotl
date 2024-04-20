package cn.toutatis.xvoid.axolotl.excel.writer.components.configuration;

import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.*;

@Data
public class AxolotlCellStyle {

    /**
     * 前景色
     */
    private AxolotlColor foregroundColor;

    /**
     * 填充模式
     */
    private FillPatternType fillPatternType;

    /**
     * 单元格水平对齐方式
     */
    private HorizontalAlignment horizontalAlignment;

    /**
     * 单元格对齐方式
     */
    private VerticalAlignment verticalAlignment;

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
     * 字体
     */
    private String fontName;

    /**
     * 字体大小
     */
    private Short fontSize;

    /**
     * 字体颜色
     */
    private IndexedColors fontColor;

    /**
     * 字体是否加粗
     */
    private Boolean fontBold;

    /**
     *设置文字为斜体
     */
    private Boolean italic;

    /**
     * 使用水平删除线
     */
    private Boolean strikeout;

    public AxolotlCellStyle foregroundColor(AxolotlColor foregroundColor) {
        this.foregroundColor = foregroundColor;
        return this;
    }

    public AxolotlCellStyle fillPatternType(FillPatternType fillPatternType) {
        this.fillPatternType = fillPatternType;
        return this;
    }

    public AxolotlCellStyle horizontalAlignment(HorizontalAlignment horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
        return this;
    }

    public AxolotlCellStyle verticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return this;
    }

    public AxolotlCellStyle borderTopStyle(BorderStyle borderTopStyle) {
        this.borderTopStyle = borderTopStyle;
        return this;
    }

    public AxolotlCellStyle topBorderColor(IndexedColors topBorderColor) {
        this.topBorderColor = topBorderColor;
        return this;
    }

    public AxolotlCellStyle borderBottomStyle(BorderStyle borderBottomStyle) {
        this.borderBottomStyle = borderBottomStyle;
        return this;
    }

    public AxolotlCellStyle bottomBorderColor(IndexedColors bottomBorderColor) {
        this.bottomBorderColor = bottomBorderColor;
        return this;
    }

    public AxolotlCellStyle borderLeftStyle(BorderStyle borderLeftStyle) {
        this.borderLeftStyle = borderLeftStyle;
        return this;
    }

    public AxolotlCellStyle leftBorderColor(IndexedColors leftBorderColor) {
        this.leftBorderColor = leftBorderColor;
        return this;
    }

    public AxolotlCellStyle borderRightStyle(BorderStyle borderRightStyle) {
        this.borderRightStyle = borderRightStyle;
        return this;
    }

    public AxolotlCellStyle rightBorderColor(IndexedColors rightBorderColor) {
        this.rightBorderColor = rightBorderColor;
        return this;
    }

    public AxolotlCellStyle fontName(String fontName) {
        this.fontName = fontName;
        return this;
    }

    public AxolotlCellStyle fontSize(Short fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    public AxolotlCellStyle fontColor(IndexedColors fontColor) {
        this.fontColor = fontColor;
        return this;
    }

    public AxolotlCellStyle fontBold(Boolean fontBold) {
        this.fontBold = fontBold;
        return this;
    }

    public AxolotlCellStyle italic(Boolean italic) {
        this.italic = italic;
        return this;
    }

    public AxolotlCellStyle strikeout(Boolean strikeout) {
        this.strikeout = strikeout;
        return this;
    }
}

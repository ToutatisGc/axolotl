package cn.toutatis.xvoid.axolotl.excel.writer.components.configuration;

import lombok.Data;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 单元格字体 配置
 * @author 张智凯
 * @version 1.0
 * 2024/3/29 0:47
 */
@Data
public class AxolotlCellFont {

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

    public AxolotlCellFont fontName(String fontName) {
        this.fontName = fontName;
        return this;
    }

    public AxolotlCellFont bold(Boolean bold) {
        this.bold = bold;
        return this;
    }

    public AxolotlCellFont fontSize(Short fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    public AxolotlCellFont fontColor(IndexedColors fontColor) {
        this.fontColor = fontColor;
        return this;
    }

    public AxolotlCellFont italic(Boolean italic) {
        this.italic = italic;
        return this;
    }

    public AxolotlCellFont strikeout(Boolean strikeout) {
        this.strikeout = strikeout;
        return this;
    }
}

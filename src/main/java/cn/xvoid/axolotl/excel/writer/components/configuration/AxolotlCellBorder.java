package cn.xvoid.axolotl.excel.writer.components.configuration;

import lombok.Data;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * 单元格边框 配置
 * @author 张智凯
 * @version 1.0
 * 2024/3/29 0:47
 */
@Data
public class AxolotlCellBorder {
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

    public AxolotlCellBorder baseBorderStyle(BorderStyle baseBorderStyle) {
        this.baseBorderStyle = baseBorderStyle;
        return this;
    }

    public AxolotlCellBorder baseBorderColor(IndexedColors baseBorderColor) {
        this.baseBorderColor = baseBorderColor;
        return this;
    }

    public AxolotlCellBorder borderTopStyle(BorderStyle borderTopStyle) {
        this.borderTopStyle = borderTopStyle;
        return this;
    }

    public AxolotlCellBorder topBorderColor(IndexedColors topBorderColor) {
        this.topBorderColor = topBorderColor;
        return this;
    }

    public AxolotlCellBorder borderBottomStyle(BorderStyle borderBottomStyle) {
        this.borderBottomStyle = borderBottomStyle;
        return this;
    }

    public AxolotlCellBorder bottomBorderColor(IndexedColors bottomBorderColor) {
        this.bottomBorderColor = bottomBorderColor;
        return this;
    }

    public AxolotlCellBorder borderLeftStyle(BorderStyle borderLeftStyle) {
        this.borderLeftStyle = borderLeftStyle;
        return this;
    }

    public AxolotlCellBorder leftBorderColor(IndexedColors leftBorderColor) {
        this.leftBorderColor = leftBorderColor;
        return this;
    }

    public AxolotlCellBorder borderRightStyle(BorderStyle borderRightStyle) {
        this.borderRightStyle = borderRightStyle;
        return this;
    }

    public AxolotlCellBorder rightBorderColor(IndexedColors rightBorderColor) {
        this.rightBorderColor = rightBorderColor;
        return this;
    }
}

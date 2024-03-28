package cn.toutatis.xvoid.axolotl.excel.writer.components;

import lombok.Data;
import org.apache.poi.ss.usermodel.*;

/**
 * 样式全局配置
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/28 9:32
 */
@Data
public class BaseCellStyle {

    /**
     * 标题行高
     */
    private short titleRowHeight;

    /**
     * 表头行高
     */
    private short headerRowHeight;

    /**
     * 内容行高
     */
    private short dataRowHeight;


    /**
     * 列宽
     */
  //  private short columnWidth;



    /** 边框相关 */

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
     * 水平对齐方式
     */
    private HorizontalAlignment horizontalAlignment;

    /**
     * 垂直对齐方式
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




    /** 字体相关 */

    /**
     * 字体名称
     */
    private String fontName;

    /**
     * 是否加粗
     */
    private boolean bold;

    /**
     * 字体大小
     */
    private short fontSize;

    /**
     *字体颜色
     */
    private IndexedColors fontColor;

    /**
     *设置文字为斜体
     */
    private boolean italic;

    /**
     * 使用水平删除线
     */
    private boolean strikeout;
}

package cn.toutatis.xvoid.axolotl.excel.writer.components;

import lombok.Data;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * 单元格主要配置
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/29 0:49
 */
@Data
public class CellMain {

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
     * 边框样式
     */
    private CellBorder border;

    /**
     * 字体样式
     */
    private CellFont font;


}

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
public class BaseCellProperty {

    /**
     * 行高
     * 预制值说明：行高在标题、表头、内容、程序常用样式中的预制值为四个不同的固定数值
     */
    private Short rowHeight;

    /**
     * 列宽
     * 列宽是指单个单元格所处列的宽度，不是某一个整体的宽度
     * 标题不支持配置列宽，标题列宽交由表头和内容进行控制
     * 预制值说明：表头列宽：依据表头单元格的值经过计算得出  内容列宽：继承表头列宽，当没有表头时，指定一个固定的值作为列宽
     * 列宽配置在开启自动列宽 AUTO_CATCH_COLUMN_LENGTH 策略时无效
     * 生效问题：因为渲染次序的问题，内容会最后渲染，而列宽是影响一整列的，自然会影响到先前已渲染的表头，所以在内容样式中配置的列宽会覆盖在表头样式中配置的列宽
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

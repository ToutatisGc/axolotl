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
     * 预制值说明：行高在标题、表头、数据中的预制值为三个不同的固定数值
     */
    private Short rowHeight;

    /**
     * 列宽
     * 列宽是指单个单元格所处列的宽度，不是某一个整体的宽度
     * 标题列宽与表头、数据列宽同步，标题不支持配置列宽
     * 预制值说明：若存在表头，每列列宽的预制值由表头最后一行每个单元格的单元格值经过计算后得出; 若不存在表头，所有列的预制值都为同一个固定的数值
     * 列宽配置在开启自动列宽 AUTO_CATCH_COLUMN_LENGTH 策略时无效
     * 生效问题：因为渲染次序的问题，数据会最后渲染，而列宽是影响一整列的，自然会影响到先前已渲染的表头，所以在数据样式中配置的列宽会覆盖在表头样式中配置的列宽
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

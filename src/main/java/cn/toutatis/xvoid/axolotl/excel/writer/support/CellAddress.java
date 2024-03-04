package cn.toutatis.xvoid.axolotl.excel.writer.support;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.math.BigDecimal;

/**
 * 用于存放模板占位符单元格的行和列
 * @author Toutatis_Gc
 */
@Data
public class CellAddress {

    public CellAddress(String cellValue, int rowPosition, int columnPosition, CellStyle cellStyle) {
        this.cellValue = cellValue;
        this.rowPosition = rowPosition;
        this.columnPosition = columnPosition;
        this.cellStyle = cellStyle;
    }

    private PlaceholderType placeholderType;

    /**
     * 模板单元格的占位符
     */
    private String placeholder;

    /**
     * 模板单元格的值
     */
    private String cellValue;

    /**
     * 模板单元格的行位置
     */
    private int rowPosition;

    /**
     * 模板单元格的列位置
     */
    private int columnPosition;

    /**
     * 模板单元格的样式
     * 一般继承自模板占位符的单元格样式
     */
    private CellStyle cellStyle;

    /**
     * 模板单元格的已写入行
     */
    private int writtenRow = -1;

    /**
     * 计算值
     */
    private BigDecimal calculatedValue = new BigDecimal("-1");

    /**
     * 合并单元格的区域
     * 需要特殊处理
     */
    private CellRangeAddress mergeRegion;

    public void setRowPosition(int rowPosition) {
        writtenRow++;
        this.rowPosition = rowPosition;
    }

    /**
     * 判断是否已经写入
     */
    public boolean isInitializedWrite() {
        return writtenRow == -1;
    }

    public String replacePlaceholder(String value) {
        return cellValue.replace(this.placeholder, value);
    }

    /**
     * 判断是否是合并单元格
     * @return 是否合并单元格
     */
    public boolean isMergeCell() {
        return mergeRegion != null;
    }

}

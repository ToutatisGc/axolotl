package cn.toutatis.xvoid.axolotl.excel.writer.support;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

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

    public String replacePlaceholder(String value) {
        return cellValue.replace(this.placeholder, value);
    }

}

package cn.toutatis.xvoid.axolotl.excel.writer.support;

import lombok.Data;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.math.BigDecimal;

/**
 * 用于存放模板占位符单元格的行和列
 * @author Toutatis_Gc
 */
@Data
public class CellAddress implements Cloneable, Serializable {

    public CellAddress(String cellValue, int rowPosition, int columnPosition, CellStyle cellStyle) {
        this.cellValue = cellValue;
        this.rowPosition = rowPosition;
        this.columnPosition = columnPosition;
        this.cellStyle = cellStyle;
    }

    private PlaceholderType placeholderType;

    /**
     * 名称
     */
    private String name;

    /**
     * 模板单元格的占位符
     */
    private String placeholder;

    /**
     * 模板单元格的值
     */
    private String cellValue;

    /**
     * 占位符产生的默认值
     */
    private String defaultValue;

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
     * 该值等于0为该单元格位置占位符，大于1说明有其他占位符，只赋予值但不位移单元格
     */
    private int sameCellPlaceholder = -1;

    /**
     * [内部维护变量]
     * 非模板单元格
     */
    private Cell _nonTemplateCell;

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

    @Override
    public CellAddress clone() {
        try {
            return (CellAddress) super.clone();
        } catch (CloneNotSupportedException e) {
            throw new AssertionError();
        }
    }

    @SneakyThrows
    public CellAddress deepClone() {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ObjectOutputStream oos = new ObjectOutputStream(bos);
        oos.writeObject(this);

        ByteArrayInputStream bis = new ByteArrayInputStream(bos.toByteArray());
        ObjectInputStream ois = new ObjectInputStream(bis);
        return (CellAddress) ois.readObject();
    }
}

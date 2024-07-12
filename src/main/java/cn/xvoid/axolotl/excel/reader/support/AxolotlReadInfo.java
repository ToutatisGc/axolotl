package cn.xvoid.axolotl.excel.reader.support;

import lombok.AccessLevel;
import lombok.Getter;
import lombok.Setter;

/**
 * 当类拥有该字段时，将为实体赋予该字段
 * 可以获取到读取时该行的信息
 * @author Toutatis_Gc
 */
@Getter
@Setter(AccessLevel.PUBLIC)
public class AxolotlReadInfo {

    public AxolotlReadInfo() {}

    /**
     * 表的索引
     */
    private Integer sheetIndex;

    /**
     * 表名称
     */
    private String sheetName;

    /**
     * 行号
     */
    private Integer rowNumber;

    @Override
    public String toString() {
        return "AxolotlReadInfo{" +
                "sheetIndex=" + sheetIndex +
                ", sheetName='" + sheetName + '\'' +
                ", rowNumber=" + rowNumber +
                '}';
    }
}

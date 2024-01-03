package cn.toutatis.xvoid.axolotl.excel.support;

import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

@Data
@NoArgsConstructor
public class CellGetInfo {

    /**
     * 对象是否使用了单元格值
     */
    private boolean alreadyFillValue = false;

    /**
     * 单元格类型
     */
    private CellType cellType;

    /**
     * 单元格值
     */
    private Object cellValue = null;

    /**
     * 单元格对象只有数字格式会将单元格赋值
     */
    private Cell _cell;

    public CellGetInfo(boolean alreadyFillValue, Object cellValue) {
        this.alreadyFillValue = alreadyFillValue;
        this.cellValue = cellValue;
    }
}

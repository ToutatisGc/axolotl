package cn.toutatis.xvoid.axolotl.excel.support;

import lombok.Data;
import lombok.NoArgsConstructor;
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

    public CellGetInfo(boolean alreadyFillValue, Object cellValue) {
        this.alreadyFillValue = alreadyFillValue;
        this.cellValue = cellValue;
    }
}

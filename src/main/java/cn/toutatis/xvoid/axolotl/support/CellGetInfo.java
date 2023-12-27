package cn.toutatis.xvoid.axolotl.support;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
@AllArgsConstructor
public class CellGetInfo {

    /**
     * 对象是否使用了单元格值
     */
    private boolean useCellValue = false;

    /**
     * 单元格值
     */
    private Object cellValue = null;
}

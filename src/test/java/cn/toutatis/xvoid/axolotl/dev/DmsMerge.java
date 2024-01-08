package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.excel.annotations.ColumnBind;
import cn.toutatis.xvoid.axolotl.excel.annotations.NamingWorkSheet;
import lombok.Data;

import java.math.BigDecimal;

@Data
@NamingWorkSheet(sheetName = "资产负债表（组织类）")
public class DmsMerge {

    @ColumnBind(columnIndex = 0)
    private String asset;
    @ColumnBind(columnIndex = 1)
    private Integer rowNumber1;
    @ColumnBind(columnIndex = 2)
    private BigDecimal amount1;

    @ColumnBind(columnIndex = 2)
    private String amount1String;

    @ColumnBind(columnIndex = 3)
    private String debt;
    @ColumnBind(columnIndex = 4)
    private Integer rowNumber2;
    @ColumnBind(columnIndex = 5)
    private BigDecimal amount2;

    @ColumnBind(columnIndex = 5)
    private String amount2String;


}

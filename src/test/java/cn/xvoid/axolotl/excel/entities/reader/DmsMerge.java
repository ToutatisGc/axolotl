package cn.xvoid.axolotl.excel.entities.reader;

import cn.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import cn.xvoid.axolotl.excel.reader.annotations.NamingWorkSheet;
import cn.xvoid.axolotl.excel.reader.support.AxolotlValid;
import lombok.Data;

import javax.validation.constraints.NotNull;
import java.math.BigDecimal;

@Data
@NamingWorkSheet(sheetName = "资产负债表（组织类）")
public class DmsMerge {

    @ColumnBind(columnIndex = 0)
    @NotNull(groups = {AxolotlValid.Simple.class})
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

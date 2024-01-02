
package cn.toutatis.xvoid.axolotl.entities;

import cn.toutatis.xvoid.axolotl.excel.annotations.ColumnBind;
import cn.toutatis.xvoid.axolotl.excel.annotations.IndexWorkSheet;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;

@IndexWorkSheet
@Data
@ToString(callSuper = true)
@EqualsAndHashCode(callSuper = true)
public class IndexTest extends BaseEntity {

    @ColumnBind(cellIndex = 0)
    private String column0;
    @ColumnBind(cellIndex = 1)
    private String column1;
    @ColumnBind(cellIndex = 2)
    private String column2;
    @ColumnBind(cellIndex = 3)
    private String column3;
    @ColumnBind(cellIndex = 4)
    private String column4;
    @ColumnBind(cellIndex = 5)
    private String column5;
    @ColumnBind(cellIndex = 6)
    private String column6;
    @ColumnBind(cellIndex = 7)
    private String column7;
    @ColumnBind(cellIndex = 8)
    private String column8;
    @ColumnBind(cellIndex = 9)
    private String column9;
//    @CellBindProperty(cellIndex = 10)
    private int column10;
    
    private boolean column11;

    
}


package cn.toutatis.xvoid.axolotl.entities;

import cn.toutatis.xvoid.axolotl.excel.annotations.CellBindProperty;
import cn.toutatis.xvoid.axolotl.excel.annotations.IndexWorkSheet;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;

@IndexWorkSheet
@Data
@ToString(callSuper = true)
@EqualsAndHashCode(callSuper = true)
public class IndexTest extends BaseEntity {

    @CellBindProperty(cellIndex = 0)
    private String column0;
    @CellBindProperty(cellIndex = 1)
    private String column1;
    @CellBindProperty(cellIndex = 2)
    private String column2;
    @CellBindProperty(cellIndex = 3)
    private String column3;
    @CellBindProperty(cellIndex = 4)
    private String column4;
    @CellBindProperty(cellIndex = 5)
    private String column5;
    @CellBindProperty(cellIndex = 6)
    private String column6;
    @CellBindProperty(cellIndex = 7)
    private String column7;
    @CellBindProperty(cellIndex = 8)
    private String column8;
    @CellBindProperty(cellIndex = 9)
    private String column9;
//    @CellBindProperty(cellIndex = 10)
    private int column10;
    
    private boolean column11;

    
}

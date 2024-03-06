package cn.toutatis.xvoid.axolotl.excel.entities.reader;

import cn.toutatis.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import lombok.Data;
import lombok.EqualsAndHashCode;

@EqualsAndHashCode(callSuper = true)
@Data
public class OneFieldString3Entity extends OneFieldString3EntityParent{

    @ColumnBind(columnIndex = 0)
    private String column1;

    @ColumnBind(columnIndex = 1)
    private String column2;
//
//    @ColumnBind(columnIndex = 2)
//    private String column3;


}

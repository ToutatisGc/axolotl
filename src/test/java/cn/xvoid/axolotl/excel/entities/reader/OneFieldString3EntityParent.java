package cn.xvoid.axolotl.excel.entities.reader;

import cn.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import lombok.Data;

@Data
public class OneFieldString3EntityParent {


    @ColumnBind(columnIndex = 2)
    private String column3;


}

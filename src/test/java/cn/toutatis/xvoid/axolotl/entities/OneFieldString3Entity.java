package cn.toutatis.xvoid.axolotl.entities;

import cn.toutatis.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import lombok.Data;

@Data
public class OneFieldString3Entity {

    @ColumnBind(columnIndex = 0)
    private String column1;

    @ColumnBind(columnIndex = 1)
    private String column2;

    @ColumnBind(columnIndex = 2)
    private String column3;


}

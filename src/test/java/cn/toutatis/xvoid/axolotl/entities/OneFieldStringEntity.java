package cn.toutatis.xvoid.axolotl.entities;

import cn.toutatis.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import lombok.Data;

@Data
public class OneFieldStringEntity {

    @ColumnBind(columnIndex = 0)
    private String column1;

    @ColumnBind(columnIndex = 1)
    private String column2;

    @ColumnBind(columnIndex = 2)
    private String column3;

    @ColumnBind(columnIndex = 3)
    private String column4;

    @ColumnBind(columnIndex = 4)
    private String column5;

    @ColumnBind(columnIndex = 5)
    private String column6;

    @ColumnBind(columnIndex = 6)
    private String column7;

    @ColumnBind(columnIndex = 7)
    private String column8;

    @ColumnBind(columnIndex = 8)
    private String column9;

    @ColumnBind(columnIndex = 9)
    private String column10;

    @ColumnBind(columnIndex = 10)
    private String column11;

    @ColumnBind(columnIndex = 11)
    private String column12;

}

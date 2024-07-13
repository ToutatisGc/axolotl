package cn.xvoid.axolotl.excel.entities.reader;

import cn.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import cn.xvoid.axolotl.excel.reader.annotations.KeepIntact;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import lombok.Data;

@Data
public class OneFieldStringKeepIntactEntity {

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
    @KeepIntact(excludePolicies = ExcelReadPolicy.CAST_NUMBER_TO_DATE)
    private String column10;

    @ColumnBind(columnIndex = 10)
    private String column11;

    @ColumnBind(columnIndex = 11)
    private String column12;

}

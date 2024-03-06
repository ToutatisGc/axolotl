package cn.toutatis.xvoid.axolotl.excel.entities.reader;

import cn.toutatis.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import lombok.Data;

@Data
public class HeaderTestEntity {

    @ColumnBind(headerName = "姓名")
    private String name;

    @ColumnBind(headerName = "性别")
    private String age;

    @ColumnBind(headerName = "行次")
    private String line1;
    @ColumnBind(headerName = "行次")
    private String line2;

}

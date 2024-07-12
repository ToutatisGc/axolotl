package cn.xvoid.axolotl.excel.entities.reader;

import cn.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import cn.xvoid.axolotl.excel.reader.support.AxolotlReadInfo;
import lombok.Data;
import lombok.EqualsAndHashCode;
import org.hibernate.validator.constraints.Length;

@EqualsAndHashCode(callSuper = true)
@Data
public class OneFieldString3Entity extends OneFieldString3EntityParent{

    @ColumnBind(columnIndex = 0)
    @Length(max = 3,message = "最大不能超过三位")
    private String column1;

    @ColumnBind(columnIndex = 1)
    private String column2;
//
//    @ColumnBind(columnIndex = 2)
//    private String column3;
    private AxolotlReadInfo readInfo;


}

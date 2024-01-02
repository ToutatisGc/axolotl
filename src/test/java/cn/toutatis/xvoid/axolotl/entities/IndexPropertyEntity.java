package cn.toutatis.xvoid.axolotl.entities;

import cn.toutatis.xvoid.axolotl.excel.annotations.ColumnBind;
import cn.toutatis.xvoid.axolotl.excel.annotations.IndexWorkSheet;
import cn.toutatis.xvoid.axolotl.excel.annotations.SpecifyPositionBind;
import cn.toutatis.xvoid.toolkit.constant.Time;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;

import java.time.LocalDateTime;
import java.util.Date;

@IndexWorkSheet
@Data
@ToString(callSuper = true)
@EqualsAndHashCode(callSuper = true)
public class IndexPropertyEntity extends BaseEntity {

    @SpecifyPositionBind("A5")
    private String title;

    @ColumnBind(cellIndex = 0,format = Time.HMS_COLON_FORMAT_REGEX)
    private String name;

    @ColumnBind(cellIndex = 1)
    private String age;

    private LocalDateTime date1;
    
    private Date date2;
    
}

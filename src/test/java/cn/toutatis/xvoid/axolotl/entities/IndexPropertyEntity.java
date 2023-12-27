package cn.toutatis.xvoid.axolotl.entities;

import cn.toutatis.xvoid.axolotl.annotations.CellBindProperty;
import cn.toutatis.xvoid.axolotl.annotations.IndexWorkSheet;
import cn.toutatis.xvoid.axolotl.annotations.SpecifyCellPosition;
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

    @SpecifyCellPosition("A5")
    private String title;

    @CellBindProperty(cellIndex = 0,dateFormat = Time.HMS_COLON_FORMAT_REGEX)
    private String name;

    @CellBindProperty(cellIndex = 1)
    private String age;

    private LocalDateTime date1;
    
    private Date date2;
    
}

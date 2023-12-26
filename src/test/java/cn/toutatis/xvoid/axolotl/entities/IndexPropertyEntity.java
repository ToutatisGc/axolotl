package cn.toutatis.xvoid.axolotl.entities;

import cn.toutatis.xvoid.axolotl.annotations.CellBindProperty;
import cn.toutatis.xvoid.axolotl.annotations.IndexWorkSheet;
import cn.toutatis.xvoid.axolotl.annotations.SpecifyCellPosition;
import cn.toutatis.xvoid.toolkit.constant.Time;

import java.time.LocalDateTime;
import java.util.Date;

@IndexWorkSheet
public class IndexPropertyEntity extends BaseEntity {

    @SpecifyCellPosition("A5")
    private String title;

    @CellBindProperty(cellIndex = 0,dateFormat = Time.HMS_COLON_FORMAT_REGEX)
    private String name;
    
    private String age;
    
    private LocalDateTime date1;
    
    private Date date2;
    
}

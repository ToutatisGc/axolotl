package cn.toutatis.xvoid.axolotl.entities;

import cn.toutatis.xvoid.axolotl.annotations.CellBindProperty;
import cn.toutatis.xvoid.axolotl.annotations.IndexWorkSheet;

import java.time.LocalDateTime;
import java.util.Date;

@IndexWorkSheet
public class IndexPropertyEntity {
    
    @CellBindProperty(cellIndex = 0)
    private String name;
    
    private String age;
    
    private LocalDateTime date1;
    
    private Date date2;
    
}

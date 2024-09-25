package cn.xvoid.axolotl.excel.entities.writer;

import cn.xvoid.axolotl.excel.writer.components.annotations.SheetHeader;
import lombok.Data;

@Data
public class SecurityQuestion {

    @SheetHeader(name = "问题")
    private String question;

    @SheetHeader(name = "存在问题")
    private String exist;

}

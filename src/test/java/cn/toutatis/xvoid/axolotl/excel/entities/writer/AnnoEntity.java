package cn.toutatis.xvoid.axolotl.excel.entities.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.components.annotations.SheetHeader;
import cn.toutatis.xvoid.axolotl.excel.writer.components.annotations.SheetTitle;
import lombok.Data;

@Data
@SheetTitle("测试表")
public class AnnoEntity {

    @SheetHeader(name="成员名称")
    private String name;

    @SheetHeader(name="地址")
    private String address;
}

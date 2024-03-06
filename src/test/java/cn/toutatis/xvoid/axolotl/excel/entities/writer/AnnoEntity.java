package cn.toutatis.xvoid.axolotl.excel.entities.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.components.SheetSimpleHeader;
import cn.toutatis.xvoid.axolotl.excel.writer.components.SheetTitle;
import lombok.Data;

@Data
@SheetTitle("测试表")
public class AnnoEntity {

    @SheetSimpleHeader(name="成员名称")
    private String name;

    @SheetSimpleHeader(name="地址")
    private String address;
}

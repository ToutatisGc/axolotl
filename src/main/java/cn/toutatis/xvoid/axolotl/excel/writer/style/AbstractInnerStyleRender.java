package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import lombok.Getter;
import lombok.Setter;

public abstract class AbstractInnerStyleRender implements ExcelStyleRender{

    @Setter @Getter
    protected AutoWriteConfig writeConfig;


}

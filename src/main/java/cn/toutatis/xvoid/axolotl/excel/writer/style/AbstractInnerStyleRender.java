package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.WriterConfig;
import lombok.Getter;
import lombok.Setter;

public abstract class AbstractInnerStyleRender implements ExcelStyleRender{

    @Setter @Getter
    protected WriterConfig writerConfig;


}

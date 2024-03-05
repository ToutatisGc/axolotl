package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AutoWriteContext;
import lombok.Getter;
import lombok.Setter;

/**
 * 样式渲染器抽象类
 * 继承此抽象类可以获取环境变量实现自定义样式渲染
 * @author Toutatis_Gc
 */
public abstract class AbstractStyleRender implements ExcelStyleRender{

    @Setter @Getter
    protected AutoWriteConfig writeConfig;

    @Setter @Getter
    protected AutoWriteContext context;

}

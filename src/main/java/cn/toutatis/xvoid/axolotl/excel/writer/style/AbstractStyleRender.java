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
@Getter
public abstract class AbstractStyleRender implements ExcelStyleRender{

    @Setter
    protected AutoWriteConfig writeConfig;

    @Setter
    protected AutoWriteContext context;

    /**
     * 是否是第一批次数据
     * @return true/false
     */
    public boolean isFirstBatch(){
        return context.isFirstBatch(context.getSwitchSheetIndex());
    }

}

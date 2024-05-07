package cn.toutatis.xvoid.axolotl.excel.writer.support;

import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.Header;
import lombok.Data;

import java.util.List;

/**
 * 工作表写入封装数据
 * @see cn.toutatis.xvoid.axolotl.AxolotlFaster 快捷操作类
 * @author 张智凯
 * @version 1.0.15
 */
@Data
public class SheetDataPackage {

    /**
     * 表头
     */
    private List<Header> headers;

    /**
     * 数据
     */
    private List<?> data;

    /**
     * 自动写入配置
     */
    private AutoWriteConfig autoWriteConfig;
}

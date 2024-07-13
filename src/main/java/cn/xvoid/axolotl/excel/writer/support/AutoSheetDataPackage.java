package cn.xvoid.axolotl.excel.writer.support;

import cn.xvoid.axolotl.AxolotlFaster;
import cn.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.xvoid.axolotl.excel.writer.components.widgets.Header;
import lombok.Data;

import java.util.List;

/**
 * 工作表写入封装数据
 * @see AxolotlFaster 快捷操作类
 * @author 张智凯
 * @version 1.0.15
 */
@Data
public class AutoSheetDataPackage {

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

    public AutoSheetDataPackage headers(List<Header> headers) {
        this.headers = headers;
        return this;
    }

    public AutoSheetDataPackage data(List<?> data) {
        this.data = data;
        return this;
    }

    public AutoSheetDataPackage autoWriteConfig(AutoWriteConfig autoWriteConfig) {
        this.autoWriteConfig = autoWriteConfig;
        return this;
    }
}

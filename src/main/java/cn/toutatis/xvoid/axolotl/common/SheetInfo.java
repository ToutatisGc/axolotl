package cn.toutatis.xvoid.axolotl.common;

import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.Header;
import lombok.Data;

import java.util.List;

/**
 * @author 张智凯
 * @version 1.0
 * @data 2024/4/26 16:27
 */
@Data
public class SheetInfo {

    private List<Header> headers;

    private List<?> data;

    private AutoWriteConfig autoWriteConfig;
}

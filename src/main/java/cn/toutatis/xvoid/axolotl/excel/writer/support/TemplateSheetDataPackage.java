package cn.toutatis.xvoid.axolotl.excel.writer.support;

import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.TemplateWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.Header;
import lombok.Data;

import java.io.File;
import java.util.List;
import java.util.Map;

/**
 * 工作表写入封装数据
 * @see cn.toutatis.xvoid.axolotl.AxolotlFaster 快捷操作类
 * @author 张智凯
 * @version 1.0.15
 */
@Data
public class TemplateSheetDataPackage {

    /**
     * 引用字段 ${}占位符
     */
    private Map<String,?> fixMapping;

    /**
     * 数据 #{}占位符
     */
    private List<?> datas;

    /**
     * 模板写入配置
     */
    private TemplateWriteConfig templateWriteConfig;

}

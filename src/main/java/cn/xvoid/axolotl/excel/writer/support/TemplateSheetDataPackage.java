package cn.xvoid.axolotl.excel.writer.support;

import cn.xvoid.axolotl.excel.writer.TemplateWriteConfig;
import cn.xvoid.axolotl.AxolotlFaster;
import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * 工作表写入封装数据
 * @see AxolotlFaster 快捷操作类
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

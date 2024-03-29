package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.components.CellMain;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.CommonWriteConfig;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.List;

/**
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/28 11:22
 */
public interface CellStyleConfigur {

    /**
     * 配置全局样式
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。
     * 全局样式配置优先级 AutoWriteConfig内样式有关配置 > 此处配置 > 预制样式
     */
    void globalStyleConfig(CellMain cell);

    /**
     * 配置表头样式
     * 渲染器渲染表头时调用
     * 表头样式配置优先级   Header对象内配置 > 此处配置 > 全局样式
     */
    void headerStyleConfig(CellMain cell);

    /**
     * 配置标题样式
     * 渲染器渲染表头时调用
     * 标题样式配置优先级  此处配置 > 表头样式
     */
    void titleStyleConfig(CellMain cell);

    /**
     * 配置数据样式
     *
     * @param sheet        {@link SXSSFSheet} 表示工作表对象，用于设置数据样式。
     * @param data {@link CommonWriteConfig} 表示 Excel 写入器的配置，用于根据需要进行更多的样式定制。
     */
    AxolotlWriteResult dataStyleConfig(CellMain cell, AbstractStyleRender.FieldInfo fieldInfo);

    /**
     * 在渲染完成后，调用该方法。
     * 在Close()方法中调用。
     */
    AxolotlWriteResult finish(SXSSFSheet sheet);
}

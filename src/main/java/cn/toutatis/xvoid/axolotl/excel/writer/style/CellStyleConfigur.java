package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.components.BaseCellStyle;
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
     * 初始化。
     * 多次写入时，该方法只会被调用一次。
     * 用于创建全局样式等。
     */
    BaseCellStyle globalCellStyle();

    /**
     * 渲染 Excel 表头的样式。
     *
     * @param sheet {@link SXSSFSheet} 表示工作表对象，用于设置表头样式。
     */
    AxolotlWriteResult renderHeader(SXSSFSheet sheet);

    /**
     * 渲染 Excel 数据的样式。
     *
     * @param sheet        {@link SXSSFSheet} 表示工作表对象，用于设置数据样式。
     * @param data {@link CommonWriteConfig} 表示 Excel 写入器的配置，用于根据需要进行更多的样式定制。
     */
    AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data);

    /**
     * 在渲染完成后，调用该方法。
     * 在Close()方法中调用。
     */
    AxolotlWriteResult finish(SXSSFSheet sheet);
}

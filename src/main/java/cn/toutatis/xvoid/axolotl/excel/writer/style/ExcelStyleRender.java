package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.CommonWriteConfig;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.List;

/**
 * Excel样式渲染器接口。
 * <p>
 * 该接口定义了用于渲染 Excel 表头和数据的样式的方法。
 * </p>
 * <p>
 * 实现此接口的类应该提供适当的实现，以便在生成 Excel 时能够定制表头和数据的样式。
 * </p>
 * <p>
 * Excel 样式渲染器主要用于 {@link SXSSFSheet}，这是 Apache POI 中用于支持大数据量的一种工作表类型。
 * </p>
 *
 * @author Toutatis_Gc
 */
public interface ExcelStyleRender {

    /**
     * 初始化。
     * 多次写入时，该方法只会被调用一次。
     * 可以用于创建全局样式等。
     * @param sheet {@link SXSSFSheet} 表示工作表对象，用于设置表头样式。
     */
    AxolotlWriteResult init(SXSSFSheet sheet);

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
    AxolotlWriteResult finish();
}

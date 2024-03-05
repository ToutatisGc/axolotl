package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.support.CommonWriteConfig;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.List;
import java.util.Map;

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
     * 渲染 Excel 表头的样式。
     *
     * @param sheet {@link SXSSFSheet} 表示工作表对象，用于设置表头样式。
     */
    void renderHeader(SXSSFSheet sheet);

    /**
     * 渲染 Excel 数据的样式。
     *
     * @param sheet        {@link SXSSFSheet} 表示工作表对象，用于设置数据样式。
     * @param data {@link CommonWriteConfig} 表示 Excel 写入器的配置，用于根据需要进行更多的样式定制。
     */
    void renderData(SXSSFSheet sheet, List<?> data);

    /**
     * 数据处理
     * @param sheet 工作表
     * @param singleMap 单行数据
     * @param data 列表数据
     */
    void dataProcessing(SXSSFSheet sheet, Map<String,?> singleMap, List<?> data);
}

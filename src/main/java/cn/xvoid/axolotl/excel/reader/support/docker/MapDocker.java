package cn.xvoid.axolotl.excel.reader.support.docker;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.support.CellGetInfo;


/**
 * 定义用于数据映射的接口，提供了转换和配置相关功能
 *
 * @author Toutatis_Gc
 * @param <T> 转换后的数据类型
 */
public interface MapDocker<T> {

    /**
     * 获取map拓展后缀
     * 用于标识或扩展用途，具体场景可根据后缀处理数据
     *
     * @return 后缀字符串
     */
    String getSuffix();

    /**
     * 设置转换后为null是否显示
     * 控制在数据转换过程中，是否将null值展示出来
     *
     * @param display true表示显示null值，false表示不显示
     */
    void setNullDisplay(Boolean display);

    /**
     * 获取转换后为null是否显示的设置
     * 查询当前对于null值的显示策略
     *
     * @return true如果null值显示，否则false
     */
    Boolean getNullDisplay();

    /**
     * 将单元格值转换为对应类型
     * 根据提供的单元格信息和读取配置，将一个单元格的值转换为目标类型T
     *
     * @param index 单元格的索引位置
     * @param cellGetInfo 单元格值及其相关的信息
     * @param readerConfig 读取配置信息，可能影响转换过程
     * @return 转换后的数据，类型为泛型T
     */
    T convert(int index, CellGetInfo cellGetInfo, ReaderConfig<?> readerConfig);

}

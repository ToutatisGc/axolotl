package cn.xvoid.axolotl.excel.reader.support.docker;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.support.CellGetInfo;


public interface MapDocker<T> {
//
//    /**
//     * 设置map拓展后缀
//     * @param name 后缀
//     */
//    void setSuffix(String name);
//
//    /**
//     * 获取map拓展后缀
//     * @return 后缀
//     */
//    String getSuffix();

    /**
     * 设置转换后为null是否显示
     * @param display 是否显示
     */
    void setNullDisplay(Boolean display);

    /**
     * 获取转换后为null是否显示
     * @return 是否显示
     */
    Boolean getNullDisplay();

    /**
     * 将单元格值转换为对应类型
     * @param cellGetInfo 单元格值
     * @return 转换
     */
    public T convert(int index,CellGetInfo cellGetInfo, ReaderConfig<?> readerConfig);

}

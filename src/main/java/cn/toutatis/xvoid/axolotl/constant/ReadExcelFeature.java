package cn.toutatis.xvoid.axolotl.constant;

/**
 * 读取Excel的功能可配置特性
 */
public enum ReadExcelFeature {


    /*表相关配置*/

    /**
     * 忽略空sheet的错误
     */
    IGNORE_EMPTY_SHEET_ERROR,

    /**
     * 读取的sheet数据按顺序排列
     * 在使用Map接收时，使用LinkedHashMap
     */
    SORTED_READ_SHEET_DATA,

    /**
     * 使用Map接收数据时，打印调试信息
     */
    USE_MAP_DEBUG,

    /*行数据配置*/

    /**
     * 空行也视为有效数据
     */
    INCLUDE_EMPTY_ROW

}

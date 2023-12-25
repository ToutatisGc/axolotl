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


    /*行数据配置*/

    /**
     * 空行也视为有效数据
     */
    INCLUDE_EMPTY_ROW,

    /**
     * 读取的sheet数据按顺序排列
     * 在使用Map接收时，使用LinkedHashMap
     */
    SORTED_READ_SHEET_DATA,

    /**
     * 判断数字为日期类型转换为日期格式
     */
    CAST_NUMBER_TO_DATE,

    /**
     * 若未指定此特性,在按行读取时,若没有指定列名,将不会绑定对象属性
     */
    DATA_BIND_PRECISE_LOCALIZATION,

    /**
     * 修整单元格去掉单元格左右的空格
     * 和换行符
     */
    TRIM_CELL_VALUE,

    /**
     * 使用Map接收数据时，打印调试信息
     */
    USE_MAP_DEBUG

}

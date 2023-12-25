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
    INCLUDE_EMPTY_ROW

}

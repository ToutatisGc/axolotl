package cn.toutatis.xvoid.axolotl.excel.writer.annotations;

/**
 * Excel写入标题信息
 */
public @interface ExcelTitle {

    /**
     * 自动生成时指定标题
     * @return 表头
     */
   String value();

    /**
     * 指定Excel工作表名称
     * @return 工作表名称
     */
    String sheetName();

}

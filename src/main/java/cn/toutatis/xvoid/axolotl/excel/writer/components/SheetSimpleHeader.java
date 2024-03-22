package cn.toutatis.xvoid.axolotl.excel.writer.components;

/**
 * 简单表头
 * 当使用实体写入时，该注解将作为表头显示
 * 若写入时指定了headers，该注解将被忽略
 * @author Toutatis_Gc
 */
public @interface SheetSimpleHeader {

    /**
     * 表头名称
     */
    String name();

    /**
     * 列宽
     */
    int width() default -1;

}

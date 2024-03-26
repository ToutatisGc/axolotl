package cn.toutatis.xvoid.axolotl.excel.writer.support;

import lombok.Getter;

/**
 * Excel写入策略
 * @author Toutatis_Gc
 */
@Getter
public enum ExcelWritePolicy {

    /**
     * 自动计算列长度
     */
    AUTO_CATCH_COLUMN_LENGTH(Type.BOOLEAN, true, false),

    /**
     * 自动在第一列插入编号
     */
    AUTO_INSERT_SERIAL_NUMBER(Type.BOOLEAN, true, false),

    /**
     * 默认填充单元格为白色
     */
    AUTO_FILL_DEFAULT_CELL_WHITE(Type.BOOLEAN, true, true),

    /**
     * 数据写入时，自动将数据写入到下一行
     * 不会影响原有模板数据的位置
     */
    TEMPLATE_SHIFT_WRITE_ROW(Type.BOOLEAN, true, true),

    /**
     * 为没有指定的占位符填充默认值
     */
    TEMPLATE_PLACEHOLDER_FILL_DEFAULT(Type.BOOLEAN, true, true),

    /**
     * 空值是否使用模板填充
     * true:使用模板填充
     * false:填充为空单元格
     */
    TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL(Type.BOOLEAN, true, true),

    /**
     * 非模板单元格是否模板填充
     * true:填充为单元格同样数据
     * false:该列为空
     */
    TEMPLATE_NON_TEMPLATE_CELL_FILL(Type.BOOLEAN, true, true),

    /**
     * 是否抛出异常
     * true:返回写入结果
     * false:抛出异常
     */
    SIMPLE_EXCEPTION_RETURN_RESULT(Type.BOOLEAN, true, true),
    ;

    /**
     * 写入策略类型
     */
    public enum Type {

        /**
         * 布尔型策略
         */
        BOOLEAN
    }

    /**
     * 写入策略类型
     */
    private final Type type;

    /**
     * 是否为默认策略
     */
    private final boolean defaultPolicy;

    /**
     * 策略值
     */
    private final Object value;

    ExcelWritePolicy(Type type, boolean defaultPolicy, Object value) {
        this.type = type;
        this.defaultPolicy = defaultPolicy;
        this.value = value;
    }
}

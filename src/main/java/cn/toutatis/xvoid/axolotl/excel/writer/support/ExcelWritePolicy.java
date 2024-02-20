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
    AUTO_CATCH_COLUMN_LENGTH(Type.BOOLEAN, true, true),

    /**
     * 自动在第一列插入编号
     */
    AUTO_INSERT_SERIAL_NUMBER(Type.BOOLEAN, true, false),

    /**
     * 将数据写入时，自动将数据写入到下一行
     * 不会影响原有模板数据的位置
     */
    SHIFT_WRITE_ROW(Type.BOOLEAN, true, true);

    /**
     * 写入策略类型
     */
    private enum Type {

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

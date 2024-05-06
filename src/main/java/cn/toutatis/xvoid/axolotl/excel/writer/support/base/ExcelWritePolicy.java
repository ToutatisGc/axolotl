package cn.toutatis.xvoid.axolotl.excel.writer.support.base;

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
     * 自动在结尾插入合计行
     */
    AUTO_INSERT_TOTAL_IN_ENDING(Type.BOOLEAN, true, true),

    /**
     * 默认隐藏工作表空白列
     * 注意：将影响导出性能和增大存储空间
     */
    AUTO_HIDDEN_BLANK_COLUMNS(Type.BOOLEAN, false, false),

    /**
     * 处理 列表数据占位符(#{}) 时，是否在占位符下一行创建新行进行渲染，反之则直接使用下一行进行渲染（会覆盖原行的数据）
     * 数据写入时，自动将数据写入到下一行
     * 不会影响原有模板数据的位置
     */
    TEMPLATE_SHIFT_WRITE_ROW(Type.BOOLEAN, true, true),

    /**
     * 当某个占位符的参数名在写入数据的属性名中无法找到时，是否将占位符替换为空字符串，反之会保留该占位符
     * 为没有指定的占位符填充默认值
     */
    TEMPLATE_PLACEHOLDER_FILL_DEFAULT(Type.BOOLEAN, true, true),

    /**
     * 当给占位符赋值为空时，是否只将占位符替换为空字符串，反之则将占位符所在的整个单元格的值都设为空
     * 空值是否使用模板填充
     * true:使用模板填充
     * false:填充为空单元格
     */
    TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL(Type.BOOLEAN, true, true),

    /**
     * 处理 列表数据占位符(#{}) 时，与占位符同行的其他单元格的值是否在渲染新增行时予以保留，反之新增行只渲染与占位符有关的数据
     * 非模板单元格是否模板填充
     * true:填充为单元格同样数据
     * false:该列为空
     */
    TEMPLATE_NON_TEMPLATE_CELL_FILL(Type.BOOLEAN, true, true),

    /**
     * 是否抛出异常
     * true:尽最大可能写入结果，忽略异常，将以日志形式记录异常
     * false:抛出异常
     */
    SIMPLE_EXCEPTION_RETURN_RESULT(Type.BOOLEAN, true, true),

    /**
     * 是否使用getter方法
     * 取消则直接使用反射字段值
     */
    SIMPLE_USE_GETTER_METHOD(Type.BOOLEAN, true, false),

    /**
     * 是否使用字典转换
     */
    SIMPLE_USE_DICT_CODE_TRANSFER(Type.BOOLEAN, true, false)
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

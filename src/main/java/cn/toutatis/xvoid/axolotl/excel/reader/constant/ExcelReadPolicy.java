package cn.toutatis.xvoid.axolotl.excel.reader.constant;

import lombok.Getter;

/**
 * 读取Excel的功能可配置特性
 */
@Getter
public enum ExcelReadPolicy {


    /*表相关配置*/

    /**
     * <p>忽略空sheet的错误</p>
     * <p>空sheet将返回null</p>
     */
    IGNORE_EMPTY_SHEET_ERROR(Type.BOOLEAN, true,true),

    /**
     * <p>忽略空表头的错误</p>
     * <p>若取消此项配置，在没有读取到指定表头时，将抛出异常</p>
     */
    IGNORE_EMPTY_SHEET_HEADER_ERROR(Type.BOOLEAN, true,true),

    /**
     * <p>将合并的单元格展开到合并单元格的各个单元格</p>
     * <p>配置此策略，可以随意指定合并单元格的位置而读取到值</p>
     * <p>不配置此策略将只能获取合并单元格中左上角单元格才能获取到单元格值</p>
     */
    SPREAD_MERGING_REGION(Type.BOOLEAN, true,true),

    /*行数据配置*/
    /**
     * 空行也视为有效数据
     * 读取时将转换为一个空对象
     * @since 0.0.5-ALPHA 将该策略默认关闭
     */
    INCLUDE_EMPTY_ROW(Type.BOOLEAN, true,false),

    /**
     * 读取的sheet数据按顺序排列
     * 在使用Map接收时，使用LinkedHashMap
     */
    SORTED_READ_SHEET_DATA(Type.BOOLEAN, true,true),

    /**
     * 判断数字为日期类型将转换为日期格式
     */
    CAST_NUMBER_TO_DATE(Type.BOOLEAN, true,true),

    /**
     * 精确绑定属性
     * 指定此特性,在按行读取时,若没有指定列名,将不会绑定对象属性
     * 否则将按照实体字段顺序自动按照索引绑定数据
     */
    DATA_BIND_PRECISE_LOCALIZATION(Type.BOOLEAN, true,true),

    /**
     * 修整单元格去掉单元格所有的空格和换行符
     */
    TRIM_CELL_VALUE(Type.BOOLEAN, true,true),

    /**
     * 使用Map接收数据时，打印调试信息
     */
    USE_MAP_DEBUG(Type.BOOLEAN, true,true),

    /**
     * 如果字段存在值覆盖掉原值
     */
    FIELD_EXIST_OVERRIDE(Type.BOOLEAN, true,true),

    /**
     * 读取数据后校验数据
     */
    VALIDATE_READ_ROW_DATA(Type.BOOLEAN, true,true);


    public enum Type{

        /**
         * 判断类型
         */
        BOOLEAN,
        /**
         * 可执行策略
         */
        EXECUTABLE,
    }

    /**
     * 策略类型
     */
    private final Type type;

    private final boolean defaultPolicy;

    private final Object value;


    ExcelReadPolicy(Type type, boolean defaultPolicy, Object value) {
        this.type = type;
        this.defaultPolicy = defaultPolicy;
        this.value = value;
    }

    /**
     * 获取策略的值
     * @return 布尔类型的策略
     */
    public boolean getPolicyAsBoolean(){
        return (boolean)value;
    }

}

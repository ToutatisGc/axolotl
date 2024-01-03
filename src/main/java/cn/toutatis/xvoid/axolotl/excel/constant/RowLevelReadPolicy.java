package cn.toutatis.xvoid.axolotl.excel.constant;

import lombok.Getter;

/**
 * 读取Excel的功能可配置特性
 */
@Getter
public enum RowLevelReadPolicy {


    /*表相关配置*/

    /**
     * 忽略空sheet的错误
     * 空sheet将返回null
     */
    IGNORE_EMPTY_SHEET_ERROR(Type.BOOLEAN, true,true),


    /*行数据配置*/

    /**
     * 空行也视为有效数据
     * 读取时将转换为一个空对象
     */
    INCLUDE_EMPTY_ROW(Type.BOOLEAN, true,true),

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
    FIELD_EXIST_OVERRIDE(Type.BOOLEAN, true,true);

    public enum Type{
        BOOLEAN,
        EXECUTABLE,
    }

    private final Type type;

    private final boolean defaultPolicy;

    private final Object value;


    RowLevelReadPolicy(Type type, boolean defaultPolicy, Object value) {
        this.type = type;
        this.defaultPolicy = defaultPolicy;
        this.value = value;
    }

    public boolean getPolicyAsBoolean(){
        return (boolean)value;
    }

}

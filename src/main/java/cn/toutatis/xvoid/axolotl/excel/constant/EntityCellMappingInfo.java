package cn.toutatis.xvoid.axolotl.excel.constant;

import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import lombok.Data;

/**
 * 实体映射单元格读取信息
 * @author Toutatis_Gc
 */
@Data
public class EntityCellMappingInfo<FT> {

    /**
     * 字段索引
     */
    private int fieldIndex = -1;

    /**
     * 行号
     */
    private int rowPosition = -1;

    /**
     * 列号
     */
    private int columnPosition;

    /**
     * 映射类型
     */
    private MappingType mappingType = MappingType.UNKNOWN;

    /**
     * 字段名
     */
    private String fieldName;

    /**
     * 字段类型
     */
    private Class<FT> fieldType;

    /**
     * 数据转换适配器
     */
    private Class<? extends DataCastAdapter<?>> dataCastAdapter;

    /**
     * 格式转换
     */
    private String format;

    /**
     * 单元格映射类型
     */
    public enum MappingType{

        /**
         * 索引类型
         */
        INDEX,

        /**
         * 位置类型
         */
        POSITION,

        /**
         * 未知类型
         */
        UNKNOWN

    }

    public EntityCellMappingInfo(Class<FT> fieldType) {
        this.fieldType = fieldType;
    }

    /**
     * 默认值填充基本类型
     * @param value 值
     * @return 默认值填充后的值
     */
    public Object fillDefaultPrimitiveValue(Object value) {
        if (value == null) {
            if (fieldType.isPrimitive()) {
                if (fieldType == int.class) {
                    return 0;
                } else if (fieldType == long.class) {
                    return 0L;
                } else if (fieldType == double.class) {
                    return 0.0;
                } else if (fieldType == float.class) {
                    return 0.0F;
                }else if (fieldType == boolean.class) {
                    return false;
                }else if (fieldType == char.class) {
                    return '\u0000';
                }else if (fieldType == short.class) {
                    return 0;
                }else if (fieldType == byte.class) {
                    return 0;
                }
            }
        }
        return value;
    }

    public boolean fieldIsPrimitive(){
        return fieldType.isPrimitive();
    }
}
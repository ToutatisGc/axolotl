package cn.toutatis.xvoid.axolotl.constant;

import cn.toutatis.xvoid.axolotl.annotations.CellBindProperty;
import com.fasterxml.jackson.annotation.JsonIgnore;
import lombok.Data;

/**
 * 实体映射单元格读取信息
 * @author Toutatis_Gc
 */
@Data
public class EntityCellMappingInfo {

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
    private Class<?> fieldType;

    /**
     * 单元格绑定属性
     */
    @JsonIgnore
    private CellBindProperty cellBindProperty;


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

}

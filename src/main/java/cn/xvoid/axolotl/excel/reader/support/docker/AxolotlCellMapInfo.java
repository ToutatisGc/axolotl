package cn.xvoid.axolotl.excel.reader.support.docker;

import lombok.Data;
import org.apache.poi.ss.usermodel.CellType;

import java.util.Map;

/**
 * AxolotlMapInfo类用于封装读取Map数据信息
 * 它提供了存储单元格索引、原始值、单元格类型以及额外的docker值的功能
 * 读取器读取Map类型时,
 * @author Toutatis_Gc
 */
@Data
public class AxolotlCellMapInfo {

    /**
     * cellIndex存储了当前信息对象对应的单元格索引这个索引用于标识信息对象在数据结构中的位置
     */
    private int cellIndex;

    /**
     * originalValue存储了原始的单元格值这个值未经过任何处理，保持原始数据的形态
     */
    private Object originalValue;

    /**
     * cellType获取单元格的类型，这有助于理解单元格数据的含义和如何处理这些数据
     */
    private CellType cellType;

    /**
     * dockerValues是一个键值对集合，用于存储MapDocker转换后的值
     */
    private Map<String,Object> dockerValues;

}


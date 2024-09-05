package cn.xvoid.axolotl.excel.reader.support.docker;

import cn.xvoid.axolotl.exceptions.AxolotlException;
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
    private final int cellIndex;

    /**
     * originalValue存储了原始的单元格值这个值未经过任何处理，保持原始数据的形态
     */
    private final Object originalValue;

    /**
     * cellType获取单元格的类型，这有助于理解单元格数据的含义和如何处理这些数据
     */
    private final CellType cellType;

    /**
     * dockerValues是一个键值对集合，用于存储MapDocker转换后的值
     */
    private Map<String,Object> dockerValues;

    /**
     * 设置拓展字段的自定义属性值
     * 此方法用于初始化dockerValues字段，禁止对同一对象多次赋值
     *
     * @param dockerValues 包含Docker容器属性的Map，键为属性名，值为属性值
     * @throws AxolotlException 如果尝试对已赋值的字段进行二次赋值，则抛出异常
     */
    public void setDockerValues(Map<String, Object> dockerValues) {
        if (this.dockerValues == null){
            this.dockerValues = dockerValues;
        } else {
            throw new AxolotlException("字段已被赋值,禁止二次赋值");
        }
    }

    /**
     * 获取Map拓展字段的单个属性值
     *
     * @param key 属性的键名
     * @return 属性的值，如果属性不存在则返回null
     */
    public Object getDockerValue(String key) {
        if (dockerValues == null){
            return null;
        }
        return dockerValues.get(key);
    }

    /**
     * 获取Map拓展字段的指定属性，并转换为指定的类型
     *
     * @param key 属性的键名
     * @param clazz 要转换的目标类型
     * @return 转换后的属性值，如果属性不存在或转换失败则返回null
     */
    public <T> T getDockerValue(String key, Class<T> clazz) {
        if (dockerValues == null){
            return null;
        }
        Object value = dockerValues.get(key);
        if (value == null){
            return null;
        }
        return clazz.cast(value);
    }

}


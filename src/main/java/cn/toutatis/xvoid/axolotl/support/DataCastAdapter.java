package cn.toutatis.xvoid.axolotl.support;

import org.apache.poi.ss.usermodel.CellType;

/**
 * 类型转换适配器
 * @author Toutatis_Gc
 * @param <T> 需要转换的类型
 */
public interface DataCastAdapter<T> {

    /**
     * 类型转换
     * @param cellGetInfo 单元格信息
     * @param config 需要转换的类型
     * @return 转换后的值
     */
    T cast(CellGetInfo cellGetInfo, CastContext<?> config);

    /**
     * 是否支持该类型进行转换
     * @param cellType 单元格类型
     * @param clazz 需要转换的类型
     * @return 是否支持该类型进行转换
     */
    boolean support(CellType cellType, Class<?> clazz);



}

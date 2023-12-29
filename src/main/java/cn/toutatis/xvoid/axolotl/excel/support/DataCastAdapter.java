package cn.toutatis.xvoid.axolotl.excel.support;

import org.apache.poi.ss.usermodel.CellType;

/**
 * 类型转换适配器
 * @author Toutatis_Gc
 * @param <FT> 需要转换的Java类型
 */
public interface DataCastAdapter<FT> {

    /**
     * 类型转换
     * @param cellGetInfo 单元格信息
     * @param context 需要转换的类型
     * @return 转换后的值
     */
    FT cast(CellGetInfo cellGetInfo, CastContext<FT> context);

    /**
     * 是否支持该类型进行转换
     * @param cellType 单元格类型
     * @param clazz 需要转换的类型
     * @return 是否支持该类型进行转换
     */
    boolean support(CellType cellType, Class<FT> clazz);



}

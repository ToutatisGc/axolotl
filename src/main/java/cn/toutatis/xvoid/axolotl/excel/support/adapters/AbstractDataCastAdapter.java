package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import lombok.Getter;
import lombok.Setter;

/**
 * 适配器基类
 * @param <FT> 目标类型
 */
@Getter
@Setter
public abstract class AbstractDataCastAdapter<FT> implements DataCastAdapter<FT> {

    /**
     * 读取配置
     */
    private ReaderConfig<?> readerConfig;

    /**
     * 实体映射信息
     */
    private EntityCellMappingInfo<?> entityCellMappingInfo;

}

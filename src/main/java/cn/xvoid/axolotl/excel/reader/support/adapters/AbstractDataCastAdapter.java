package cn.xvoid.axolotl.excel.reader.support.adapters;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.xvoid.axolotl.excel.reader.support.DataCastAdapter;
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

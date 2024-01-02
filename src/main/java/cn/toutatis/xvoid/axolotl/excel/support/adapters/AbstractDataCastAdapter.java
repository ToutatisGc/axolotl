package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public abstract class AbstractDataCastAdapter<FT> implements DataCastAdapter<FT> {

    private ReaderConfig<?> readerConfig;

    private EntityCellMappingInfo<?> entityCellMappingInfo;

}

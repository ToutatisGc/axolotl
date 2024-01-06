package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.exceptions.AxolotlExcelReadException;
import org.apache.poi.ss.usermodel.CellType;

/**
 * 自动适配器
 */
public class AutoAdapter extends AbstractDataCastAdapter<Object> {

    /**
     * 获取自动转换器实例
     */
    public static AutoAdapter instance() {
        return new AutoAdapter();
    }
    @Override
    @SuppressWarnings({"unchecked","rawtypes"})
    public Object cast(CellGetInfo cellGetInfo, CastContext context) {
        DataCastAdapter<?> adapter = DefaultAdapters.getAdapter(context.getCastType());
        return adapter.cast(cellGetInfo,context);
    }

    @Override
    @SuppressWarnings({"unchecked","rawtypes"})
    public boolean support(CellType cellType, Class clazz) {
        DataCastAdapter<?> adapter = DefaultAdapters.getAdapter(clazz);
        EntityCellMappingInfo<?> entityCellMappingInfo = getEntityCellMappingInfo();
        if (adapter == null){
            throw new AxolotlExcelReadException(
                    entityCellMappingInfo,
                    "未找到可转换的字段类型:[%s],请配置适配器".formatted(
                            entityCellMappingInfo.getFieldType().getSimpleName()
                    )
            );
        }
        if (adapter instanceof AbstractDataCastAdapter abstractDataCastAdapter){
            abstractDataCastAdapter.setReaderConfig(getReaderConfig());
            abstractDataCastAdapter.setEntityCellMappingInfo(entityCellMappingInfo);
            return abstractDataCastAdapter.support(cellType,clazz);
        }
        return adapter.support(cellType,clazz);
    }
}

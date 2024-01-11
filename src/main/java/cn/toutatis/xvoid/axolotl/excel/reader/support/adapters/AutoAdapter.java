package cn.toutatis.xvoid.axolotl.excel.reader.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.reader.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.reader.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
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
            String msg = String.format("未找到可转换的字段类型:[%s],请配置适配器", entityCellMappingInfo.getFieldType().getSimpleName());
            throw new AxolotlExcelReadException(entityCellMappingInfo, msg);
        }
        if (adapter instanceof AbstractDataCastAdapter){
            AbstractDataCastAdapter abstractDataCastAdapter = (AbstractDataCastAdapter) adapter;
            abstractDataCastAdapter.setReaderConfig(getReaderConfig());
            abstractDataCastAdapter.setEntityCellMappingInfo(entityCellMappingInfo);
            return abstractDataCastAdapter.support(cellType,clazz);
        }
        return adapter.support(cellType,clazz);
    }
}

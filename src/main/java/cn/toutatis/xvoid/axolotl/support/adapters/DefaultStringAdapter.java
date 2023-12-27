package cn.toutatis.xvoid.axolotl.support.adapters;

import cn.toutatis.xvoid.axolotl.support.CastContext;
import cn.toutatis.xvoid.axolotl.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.support.DataCastAdapter;
import org.apache.poi.ss.usermodel.CellType;

/**
 * 默认的String类型适配器
 * @author Toutatis_Gc
 */
public class DefaultStringAdapter extends AbstractDataCastAdapter<String> implements DataCastAdapter<String> {
    @Override
    public String cast(CellGetInfo cellGetInfo, CastContext<?> config) {
        return null;
    }

    @Override
    public boolean support(CellType cellType, Class<?> clazz) {
        return true;
    }
}

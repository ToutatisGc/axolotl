package cn.toutatis.xvoid.axolotl.support.adapters;

import cn.toutatis.xvoid.axolotl.support.CastContext;
import cn.toutatis.xvoid.axolotl.support.CellGetInfo;
import org.apache.poi.ss.usermodel.CellType;

public class AutoAdapter extends AbstractDataCastAdapter<Object,Object>{
    @Override
    public Object cast(CellGetInfo cellGetInfo, CastContext<Object> context) {
        //TODO 将默认转换器放到该位置
        return null;
    }

    @Override
    public boolean support(CellType cellType, Class<Object> clazz) {
        return false;
    }
}

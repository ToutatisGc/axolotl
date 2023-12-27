package cn.toutatis.xvoid.axolotl.support.adapters;

import cn.toutatis.xvoid.axolotl.support.CastConfig;
import cn.toutatis.xvoid.axolotl.support.DataCastAdapter;
import org.apache.poi.ss.usermodel.CellType;

/**
 * 默认的String类型适配器
 * @author Toutatis_Gc
 */
public class DefaultStringAdapter extends AbstractDataCastAdapter<String> implements DataCastAdapter<String> {
    @Override
    public String cast(CellType cellType,Object value, CastConfig<String> config) {
//        Object o = this.fillDefaultPrimitiveValue(value, config.getCastType());
//        return switch (cellType) {
//            case NUMERIC -> {
//                {
//                    getWorkBookReaderConfig().getReadFeatureAsBoolean(ReadExcelFeature.CAST_NUMBER_TO_DATE);
//                }
//                yield value.toString();
//            }
//            case BLANK -> null;
//            case BOOLEAN -> value.toString();
//            default -> value.toString();
//        };
        return null;
    }

    @Override
    public boolean support(CellType cellType, Class<String> clazz) {
        return true;
    }
}

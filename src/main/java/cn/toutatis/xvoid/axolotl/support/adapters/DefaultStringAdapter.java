package cn.toutatis.xvoid.axolotl.support.adapters;

import cn.toutatis.xvoid.axolotl.constant.ReadExcelFeature;
import cn.toutatis.xvoid.axolotl.support.CastContext;
import cn.toutatis.xvoid.axolotl.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.support.WorkBookReaderConfig;
import cn.toutatis.xvoid.toolkit.constant.Regex;
import org.apache.poi.ss.usermodel.CellType;

/**
 * 默认的String类型适配器
 * @author Toutatis_Gc
 */
public class DefaultStringAdapter extends AbstractDataCastAdapter<String> implements DataCastAdapter<String> {
    @Override
    public String cast(CellGetInfo cellGetInfo, CastContext<String> config) {
        WorkBookReaderConfig<String> readerConfig =  getWorkBookReaderConfig();
        Object cellValue = cellGetInfo.getCellValue();
        if (cellValue == null){return null;}
        switch (cellGetInfo.getCellType()){
            case STRING:{
                if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.TRIM_CELL_VALUE)){
                    cellValue = Regex.convertSingleLine(cellValue.toString());
                }
                return cellValue.toString();
            }
            case BOOLEAN, NUMERIC, FORMULA:
                return cellGetInfo.getCellValue().toString();
            case BLANK:
            default:
                return null;
        }
    }

    @Override
    public boolean support(CellType cellType, Class<String> clazz) {
        return true;
    }
}

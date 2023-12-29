package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.constant.ReadExcelFeature;
import cn.toutatis.xvoid.axolotl.excel.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.WorkBookReaderConfig;
import cn.toutatis.xvoid.toolkit.constant.Regex;
import cn.toutatis.xvoid.toolkit.constant.Time;
import org.apache.poi.ss.usermodel.CellType;

import java.util.Date;

/**
 * 默认的String类型适配器
 * @author Toutatis_Gc
 */
public class DefaultStringAdapter extends AbstractDataCastAdapter<Object,String> implements DataCastAdapter<String> {
    @Override
    public String cast(CellGetInfo cellGetInfo, CastContext<String> context) {
        WorkBookReaderConfig<Object> readerConfig =  getWorkBookReaderConfig();
        Object cellValue = cellGetInfo.getCellValue();
        if (cellValue == null){return null;}
        switch (cellGetInfo.getCellType()){
            case STRING:{
                if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.TRIM_CELL_VALUE)){
                    cellValue = Regex.convertSingleLine(cellValue.toString());
                }
                return cellValue.toString();
            }
            case NUMERIC:{
                if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.CAST_NUMBER_TO_DATE)){
//                    if (DateUtil.isCellDateFormatted())
                    cellValue = Time.regexTime(context.getDataFormat(),new Date());
                }
            }
            case BOOLEAN, FORMULA:
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

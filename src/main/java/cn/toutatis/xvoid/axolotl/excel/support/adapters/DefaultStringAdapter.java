package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.constant.ReadExcelFeature;
import cn.toutatis.xvoid.axolotl.excel.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.toolkit.constant.Regex;
import cn.toutatis.xvoid.toolkit.constant.Time;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.util.Map;

/**
 * 默认的String类型适配器
 * @author Toutatis_Gc
 */
public class DefaultStringAdapter extends AbstractDataCastAdapter<String> implements DataCastAdapter<String> {
    @Override
    public String cast(CellGetInfo cellGetInfo, CastContext<String> context) {
        ReaderConfig<?> readerConfig = getReaderConfig();
        Object cellValue = cellGetInfo.getCellValue();
        if (!cellGetInfo.isUseCellValue()){
            return context.getCastType().cast(cellValue);
        }
        EntityCellMappingInfo<?> entityCellMappingInfo = getEntityCellMappingInfo();
        Map<ReadExcelFeature, Object> excelFeatures = entityCellMappingInfo.getExcelFeatures();
        return switch (cellGetInfo.getCellType()) {
            case STRING -> {
                if (!excelFeatures.containsKey(ReadExcelFeature.TRIM_CELL_VALUE)) {
                    if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.TRIM_CELL_VALUE)) {
                        cellValue = Regex.convertSingleLine(cellValue.toString());
                    }
                }
                yield cellValue.toString();
            }
            case NUMERIC -> {
                if (!excelFeatures.containsKey(ReadExcelFeature.CAST_NUMBER_TO_DATE)) {
                    if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.CAST_NUMBER_TO_DATE)) {
                        if (DateUtil.isValidExcelDate((Double) cellValue)) {
                            cellValue = Time.regexTime(context.getDataFormat(), DateUtil.getJavaDate((Double) cellValue));
                        }
                    }
                }
                yield "%s".formatted(cellValue);
            }
            case BOOLEAN, FORMULA -> cellGetInfo.getCellValue().toString();
            default -> null;
        };
    }

    @Override
    public boolean support(CellType cellType, Class<String> clazz) {
        return true;
    }
}

package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.constant.RowLevelReadPolicy;
import cn.toutatis.xvoid.axolotl.excel.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.toolkit.constant.Regex;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.validator.Validator;
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
        if (!cellGetInfo.isAlreadyFillValue()){
            return context.getCastType().cast(cellValue);
        }
        EntityCellMappingInfo<?> entityCellMappingInfo = getEntityCellMappingInfo();
        Map<RowLevelReadPolicy, Object> excludePolicies = entityCellMappingInfo.getExcludePolicies();
        return switch (cellGetInfo.getCellType()) {
            case STRING -> {
                String cellValueString = (String) cellValue;
                if (!excludePolicies.containsKey(RowLevelReadPolicy.TRIM_CELL_VALUE)) {
                    if (readerConfig.getReadPolicyAsBoolean(RowLevelReadPolicy.TRIM_CELL_VALUE)) {
                        cellValueString = Regex.convertSingleLine(cellValueString).replace(" ","");
                    }
                }
                if (Validator.strIsNumber(cellValueString)){
                    if (!excludePolicies.containsKey(RowLevelReadPolicy.CAST_NUMBER_TO_DATE)) {
                        if (readerConfig.getReadPolicyAsBoolean(RowLevelReadPolicy.CAST_NUMBER_TO_DATE)) {
                            if (DateUtil.isCellDateFormatted(cellGetInfo.get_cell())) {
                                cellValueString = Time.regexTime(context.getDataFormat(), DateUtil.getJavaDate(Double.parseDouble(cellValueString)));
                            }
                        }
                    }
                }
                yield cellValueString;
            }
            case NUMERIC -> {
                if (!excludePolicies.containsKey(RowLevelReadPolicy.CAST_NUMBER_TO_DATE)) {
                    if (readerConfig.getReadPolicyAsBoolean(RowLevelReadPolicy.CAST_NUMBER_TO_DATE)) {
                        if (DateUtil.isCellDateFormatted(cellGetInfo.get_cell())) {
                            cellValue = Time.regexTime(context.getDataFormat(), DateUtil.getJavaDate((Double) cellValue));
                        }
                    }
                }
                yield "%s".formatted(cellValue);
            }
            case BOOLEAN, FORMULA -> cellValue.toString();
            default -> null;
        };
    }

    @Override
    public boolean support(CellType cellType, Class<String> clazz) {
        return true;
    }
}

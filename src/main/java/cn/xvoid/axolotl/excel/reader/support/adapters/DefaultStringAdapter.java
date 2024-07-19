package cn.xvoid.axolotl.excel.reader.support.adapters;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.CastContext;
import cn.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.xvoid.axolotl.excel.reader.support.DataCastAdapter;
import cn.xvoid.axolotl.excel.writer.style.StyleHelper;
import cn.xvoid.common.standard.StringPool;
import cn.xvoid.toolkit.constant.Regex;
import cn.xvoid.toolkit.constant.Time;
import cn.xvoid.toolkit.validator.Validator;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.DecimalFormat;
import java.util.Map;

/**
 * 默认的String类型适配器
 * @author Toutatis_Gc
 */
public class DefaultStringAdapter extends AbstractDataCastAdapter<String> implements DataCastAdapter<String> {
    @Override
    public String cast(CellGetInfo cellGetInfo, CastContext<String> context) {
        Object cellValue = cellGetInfo.getCellValue();
        if (!cellGetInfo.isAlreadyFillValue()){
            return context.getCastType().cast(cellValue);
        }
        ReaderConfig<?> readerConfig = getReaderConfig();
        EntityCellMappingInfo<?> entityCellMappingInfo = getEntityCellMappingInfo();
        Map<ExcelReadPolicy, Object> excludePolicies = entityCellMappingInfo.getExcludePolicies();
        switch (cellGetInfo.getCellType()) {
            case STRING:
                String cellValueString = (String) cellValue;
                if (!excludePolicies.containsKey(ExcelReadPolicy.TRIM_CELL_VALUE)) {
                    if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.TRIM_CELL_VALUE)) {
                        cellValueString = Regex.convertSingleLine(cellValueString).replace(" ", "");
                    }
                }
                if (Validator.strIsNumber(cellValueString)) {
                    if (!excludePolicies.containsKey(ExcelReadPolicy.CAST_NUMBER_TO_DATE)) {
                        if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.CAST_NUMBER_TO_DATE)) {
                            if (DateUtil.isCellDateFormatted(cellGetInfo.get_cell())) {
                                cellValueString = Time.regexTime(context.getDataFormat(), DateUtil.getJavaDate(Double.parseDouble(cellValueString)));
                            }
                        }
                    }
                }
                return cellValueString;
            case NUMERIC:
                // 判断是否为文本格式
                if (!excludePolicies.containsKey(ExcelReadPolicy.CAST_NUMBER_TO_DATE)) {
                    if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.CAST_NUMBER_TO_DATE)) {
                        if (DateUtil.isCellDateFormatted(cellGetInfo.get_cell())) {
                            cellValue = Time.regexTime(context.getDataFormat(), DateUtil.getJavaDate((Double) cellValue));
                            return String.format("%s", cellValue);
                        }
                    }
                }
                short dataFormat = cellGetInfo.get_cell().getCellStyle().getDataFormat();
                boolean isText = (dataFormat == StyleHelper.DATA_FORMAT_GENERAL_INDEX || dataFormat == StyleHelper.DATA_FORMAT_PLAIN_TEXT_INDEX);
                if (isText){
                    DecimalFormat decimalFormat = new DecimalFormat(StringPool.HASH);
                    return decimalFormat.format(cellValue);
//                    if ((Double) cellValue % 1 == 0) {
//                        return Integer.toString(((Double) cellValue).intValue());
//                    }else {
////                        DecimalFormat decimalFormat = new DecimalFormat(StringPool.HASH);
////                        return decimalFormat.format(cellValue);
//                        return cellValue.toString();
//                    }
                }else {
                    if ((Double) cellValue % 1 == 0) {
                        return Integer.toString(((Double) cellValue).intValue());
                    } else {
                        return String.format("%." + AxolotlDefaultReaderConfig.XVOID_DEFAULT_DECIMAL_SCALE + "f", (Double) cellValue);
                    }
                }
            case BOOLEAN:
            case FORMULA:
                return cellValue.toString();
            default:
                return null;
        }
    }

    @Override
    public boolean support(CellType cellType, Class<String> clazz) {
        return clazz == String.class ;
    }
}

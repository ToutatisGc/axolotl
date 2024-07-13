package cn.xvoid.axolotl.excel.reader.support.adapters;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.CastContext;
import cn.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.xvoid.axolotl.excel.reader.support.DataCastAdapter;
import cn.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.xvoid.toolkit.constant.Regex;
import cn.xvoid.toolkit.validator.Validator;
import org.apache.poi.ss.usermodel.CellType;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Map;

public class DefaultNumericAdapter<NT> extends AbstractDataCastAdapter<NT> implements DataCastAdapter<NT> {

    private final Class<NT> numberClass;

    public DefaultNumericAdapter(Class<NT> numberClass) {
        this.numberClass = numberClass;
    }

    @Override
    public NT cast(CellGetInfo cellGetInfo, CastContext<NT> context) {
        Object cellValue = cellGetInfo.getCellValue();
        if (!cellGetInfo.isAlreadyFillValue()){
            return numberClass.cast(cellValue);
        }
        switch (cellGetInfo.getCellType()){
            case NUMERIC:
                Double doubleValue = (Double) cellValue;
                return this.castDoubleToOtherTypeNumber(doubleValue,context);
            case STRING:
                ReaderConfig<?> readerConfig = getReaderConfig();
                EntityCellMappingInfo<?> entityCellMappingInfo = getEntityCellMappingInfo();
                Map<ExcelReadPolicy, Object> excelPolicies = entityCellMappingInfo.getExcludePolicies();
                if (!excelPolicies.containsKey(ExcelReadPolicy.TRIM_CELL_VALUE)) {
                    if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.TRIM_CELL_VALUE)) {
                        cellValue = Regex.convertSingleLine(cellValue.toString());
                    }
                }
                if (Validator.strIsNumber((String) cellValue)){
                    Double cellDoubleValue = Double.valueOf((String) cellValue);
                    return this.castDoubleToOtherTypeNumber(cellDoubleValue,context);
                }else {
                    if("".equals(cellValue)){
                        if (!context.getCastType().isPrimitive()){return null;}
                    }
                    throw new AxolotlExcelReadException(context,"字符串不是数字格式无法转换");
                }
            case BOOLEAN:
                if ((boolean)cellGetInfo.getCellValue()){
                    return numberClass.cast(1);
                }else {
                    return numberClass.cast(0);
                }
            case BLANK:
            default:
                return null;
        }
    }

    @SuppressWarnings("unchecked")
    private NT castDoubleToOtherTypeNumber(Double doubleValue,CastContext<NT> context) {
        if (numberClass.equals(Double.class)) {
            return (NT) doubleValue;
        } else if (numberClass.equals(BigDecimal.class)){
            BigDecimal bigDecimal =
                    new BigDecimal(doubleValue.toString())
                    .setScale(AxolotlDefaultReaderConfig.XVOID_DEFAULT_DECIMAL_SCALE, RoundingMode.HALF_UP);
            return (NT) bigDecimal;
        } else if (numberClass.equals(Integer.class)) {
            return (NT) Integer.valueOf(doubleValue.intValue());
        } else if (numberClass.equals(Float.class)) {
            return (NT) Float.valueOf(doubleValue.floatValue());
        } else if (numberClass.equals(Long.class)) {
            return (NT) Long.valueOf(doubleValue.longValue());
        } else if (numberClass.equals(Short.class)) {
            return (NT) Short.valueOf(doubleValue.shortValue());
        }else {
            throw new AxolotlExcelReadException(context,"不支持的数字类型转换");
        }
    }

    @Override
    public boolean support(CellType cellType, Class<NT> clazz) {
        return (cellType == CellType.NUMERIC ||
                cellType == CellType.STRING ||
                cellType == CellType.BOOLEAN ||
                cellType == CellType.BLANK) &&
                (clazz == Integer.class || clazz == String.class ||
                clazz == int.class || clazz == BigDecimal.class ||
                clazz == Long.class || clazz == long.class ||
                clazz == Double.class || clazz == double.class ||
                clazz == Float.class || clazz == float.class ||
                clazz == Short.class || clazz == short.class ||
                clazz == Byte.class || clazz == byte.class ||
                clazz == Number.class) ;
    }
}

package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.constant.RowLevelReadPolicy;
import cn.toutatis.xvoid.axolotl.excel.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.toolkit.constant.Regex;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import org.apache.poi.ss.usermodel.CellType;

import java.math.BigDecimal;
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
                return this.castDoubleToNumber(doubleValue);
            case STRING:
                ReaderConfig<?> readerConfig = getReaderConfig();
                EntityCellMappingInfo<?> entityCellMappingInfo = getEntityCellMappingInfo();
                Map<RowLevelReadPolicy, Object> excelPolicies = entityCellMappingInfo.getExcludePolicies();
                if (!excelPolicies.containsKey(RowLevelReadPolicy.TRIM_CELL_VALUE)) {
                    if (readerConfig.getReadPolicyAsBoolean(RowLevelReadPolicy.TRIM_CELL_VALUE)) {
                        cellValue = Regex.convertSingleLine(cellValue.toString());
                    }
                }
                if (Validator.strIsNumber((String) cellValue)){
                    Double cellDoubleValue = Double.valueOf((String) cellValue);
                    return this.castDoubleToNumber(cellDoubleValue);
                }else {
                    throw new AxolotlExcelReadException("字符串不是数字格式无法转换");
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
    private NT castDoubleToNumber(Double doubleValue) {
        if (numberClass.equals(Double.class)) {
            return (NT) doubleValue;
        } else if (numberClass.equals(BigDecimal.class)){
            return (NT) new BigDecimal(doubleValue.toString());
        } else if (numberClass.equals(Integer.class)) {
            return (NT) Integer.valueOf(doubleValue.intValue());
        } else if (numberClass.equals(Float.class)) {
            return (NT) Float.valueOf(doubleValue.floatValue());
        } else if (numberClass.equals(Long.class)) {
            return (NT) Long.valueOf(doubleValue.longValue());
        } else if (numberClass.equals(Short.class)) {
            return (NT) Short.valueOf(doubleValue.shortValue());
        }else {
            throw new AxolotlExcelReadException("不支持的数字类型转换");
        }
    }

    @Override
    public boolean support(CellType cellType, Class<NT> clazz) {
        return clazz == Integer.class || clazz == String.class ||
                clazz == int.class ||
                clazz == Long.class || clazz == long.class ||
                clazz == Double.class || clazz == double.class ||
                clazz == Float.class || clazz == float.class ||
                clazz == Short.class || clazz == short.class ||
                clazz == Byte.class || clazz == byte.class ||
                clazz == Number.class ;
    }
}

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
                return numberClass.cast(cellGetInfo.getCellValue());
            case STRING:
                ReaderConfig<?> readerConfig = getReaderConfig();
                EntityCellMappingInfo<?> entityCellMappingInfo = getEntityCellMappingInfo();
                Map<RowLevelReadPolicy, Object> excelPolicies = entityCellMappingInfo.getExcelPolicies();
                if (!excelPolicies.containsKey(RowLevelReadPolicy.TRIM_CELL_VALUE)) {
                    if (readerConfig.getReadPolicyAsBoolean(RowLevelReadPolicy.TRIM_CELL_VALUE)) {
                        cellValue = Regex.convertSingleLine(cellValue.toString());
                    }
                }
                if (Validator.strIsNumber((String) cellValue)){
                    return numberClass.cast(cellGetInfo.getCellValue());
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

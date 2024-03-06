package cn.toutatis.xvoid.axolotl.excel.reader.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.reader.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.reader.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import org.apache.poi.ss.usermodel.CellType;

import java.util.HashMap;
import java.util.Map;

public class DefaultBooleanAdapter extends AbstractDataCastAdapter<Boolean> implements DataCastAdapter<Boolean> {

    /**
     * 映射表
     */
    private static final Map<Object,Boolean> trueFalseMap = new HashMap<>();

    static {
        trueFalseMap.put("TRUE", true);
        trueFalseMap.put("FALSE", false);
        trueFalseMap.put("1", true);
        trueFalseMap.put("0", false);
        trueFalseMap.put("是", true);
        trueFalseMap.put("否", false);
        trueFalseMap.put("ON", true);
        trueFalseMap.put("OFF", false);
    }

    @Override
    public Boolean cast(CellGetInfo cellGetInfo, CastContext<Boolean> context) {
        if (!cellGetInfo.isAlreadyFillValue()){
            return (Boolean) cellGetInfo.getCellValue();
        }
        switch (cellGetInfo.getCellType()){
            case BOOLEAN:
                return (Boolean) cellGetInfo.getCellValue();
            case STRING:
                String cellValue = (String) cellGetInfo.getCellValue();
                String upperCase = cellValue.toUpperCase();
                return trueFalseMap.getOrDefault(upperCase, false);
            case NUMERIC:
                if (cellGetInfo.getCellValue() instanceof Number){
                    Object cellGetInfoCellValue = cellGetInfo.getCellValue();
                    return ((Number) cellGetInfoCellValue).intValue() == 1;
                }
            default:
                break;
        }

        throw new AxolotlExcelReadException(context,String.format("无法将值[%s]转换为布尔值",cellGetInfo.getCellValue()));
    }

    @Override
    public boolean support(CellType cellType, Class<Boolean> clazz) {
        return (cellType == CellType.BOOLEAN ||
                cellType == CellType.STRING ||
                cellType == CellType.NUMERIC) &&
                (clazz == Boolean.class ||
                clazz == boolean.class) ;
    }

}

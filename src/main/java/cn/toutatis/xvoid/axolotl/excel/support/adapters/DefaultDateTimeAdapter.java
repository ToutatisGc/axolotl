package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.constant.RowLevelReadPolicy;
import cn.toutatis.xvoid.axolotl.excel.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.toolkit.constant.Regex;
import cn.toutatis.xvoid.toolkit.constant.Time;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.Map;

public class DefaultDateTimeAdapter<NT> extends AbstractDataCastAdapter<NT> implements DataCastAdapter<NT> {

    private final Class<NT> dateClass;

    public DefaultDateTimeAdapter(Class<NT> dateClass) {
        this.dateClass = dateClass;
    }


    @Override
    public NT cast(CellGetInfo cellGetInfo, CastContext<NT> context) {
        Object cellValue = cellGetInfo.getCellValue();
        if (!cellGetInfo.isAlreadyFillValue()){
            return null;
        }
        ReaderConfig<?> readerConfig = getReaderConfig();
        EntityCellMappingInfo<?> entityCellMappingInfo = getEntityCellMappingInfo();
        Map<RowLevelReadPolicy, Object> excelPolicies = entityCellMappingInfo.getExcludePolicies();
        switch (cellGetInfo.getCellType()){
            case STRING:
                if (dateClass == LocalDateTime.class){
                    if (!excelPolicies.containsKey(RowLevelReadPolicy.TRIM_CELL_VALUE)) {
                        if (readerConfig.getReadPolicyAsBoolean(RowLevelReadPolicy.TRIM_CELL_VALUE)) {
                            cellValue = Regex.convertSingleLine(cellValue.toString());
                            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(context.getDataFormat());
                            LocalDateTime localDateTime = LocalDateTime.parse((CharSequence) cellValue, formatter);
                            return dateClass.cast(localDateTime);
                        }
                    }
                }else if (dateClass == Date.class){
                    SimpleDateFormat format = new SimpleDateFormat(context.getDataFormat());
                    Date date = Time.parseData(format, cellValue.toString());
                    return dateClass.cast(date);
                }
            case NUMERIC:
                if (DateUtil.isValidExcelDate((Double) cellValue)) {
                    Date javaDate = DateUtil.getJavaDate((Double) cellValue);
                    if (dateClass == Date.class){
                        return dateClass.cast(javaDate);
                    }else if (dateClass == LocalDateTime.class){
                        Instant instant = javaDate.toInstant();
                        ZonedDateTime zonedDateTime = instant.atZone(ZoneId.systemDefault());
                        return dateClass.cast(zonedDateTime.toLocalDateTime());
                    }
                }else {
                    throw new AxolotlExcelReadException("读取值[%s]无法转换日期格式,请自行转换格式".formatted(cellValue));
                }
            default:
                throw new AxolotlExcelReadException("读取类型[%s]无法转换日期格式".formatted(cellGetInfo.getCellType()));
        }
    }

    @Override
    public boolean support(CellType cellType, Class<NT> clazz) {
        return clazz == Date.class ||
                clazz == LocalDateTime.class;
    }
}
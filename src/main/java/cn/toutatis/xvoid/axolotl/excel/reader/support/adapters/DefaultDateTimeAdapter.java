package cn.toutatis.xvoid.axolotl.excel.reader.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.toutatis.xvoid.axolotl.excel.reader.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.reader.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.toolkit.constant.Regex;
import cn.toutatis.xvoid.toolkit.constant.Time;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.text.SimpleDateFormat;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
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
        Map<ExcelReadPolicy, Object> excelPolicies = entityCellMappingInfo.getExcludePolicies();
        switch (cellGetInfo.getCellType()){
            case STRING:
                if (dateClass == LocalDateTime.class){
                    if (!excelPolicies.containsKey(ExcelReadPolicy.TRIM_CELL_VALUE)) {
                        if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.TRIM_CELL_VALUE)) {
                            cellValue = Regex.convertSingleLine(cellValue.toString());
                        }
                    }
                    if (cellValue.toString().length() != context.getDataFormat().length()){
                        throw new AxolotlExcelReadException(context,String.format("请指定正确的日期格式,获取为:[%s],尝试转换为:[%s]",cellValue, context.getDataFormat()));
                    }
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(context.getDataFormat());
                    LocalDateTime localDateTime;
                    try {
                        localDateTime = LocalDateTime.parse((CharSequence) cellValue, formatter);
                    }catch (DateTimeParseException timeParseException){
                        if (timeParseException.getMessage().contains("Unable to obtain LocalDateTime from")){
                            localDateTime = LocalDate.parse(cellValue.toString(), DateTimeFormatter.ofPattern("yyyy-MM-dd")).atStartOfDay();
                        }else {
                            throw timeParseException;
                        }
                    }
                    return dateClass.cast(localDateTime);
                }else if (dateClass == Date.class){
                    SimpleDateFormat format = new SimpleDateFormat(context.getDataFormat());
                    try {
                        Date date = Time.parseData(format, cellValue.toString());
                        return dateClass.cast(date);
                    }catch(Exception parseException){
                        throw new AxolotlExcelReadException(context,String.format("请指定正确的日期格式,获取为:[%s],尝试转换为:[%s]",cellValue, context.getDataFormat()));
                    }
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
                    throw new AxolotlExcelReadException(context,String.format("读取值[%s]无法转换日期格式,请自行转换格式",cellValue));
                }
            case BLANK:
                return null;
            default:
                throw new AxolotlExcelReadException(context,String.format("单元格位置:[%s]读取类型[%s]无法转换日期格式",context.getHumanReadablePosition(),cellGetInfo.getCellType()));
        }
    }

    @Override
    public boolean support(CellType cellType, Class<NT> clazz) {
        return clazz == Date.class ||
                clazz == LocalDateTime.class;
    }
}

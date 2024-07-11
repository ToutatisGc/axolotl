package cn.toutatis.xvoid.axolotl.excel.reader.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.toutatis.xvoid.axolotl.excel.reader.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.reader.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.xvoid.toolkit.constant.Regex;
import cn.xvoid.toolkit.constant.Time;
import cn.xvoid.toolkit.log.LoggerToolkit;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.slf4j.Logger;

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

    /**
     * 日志工具
     */
    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlExcelReader.class);


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
                    cellValue = checkTrimFeature2NewString(excelPolicies,readerConfig,context,cellValue);
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(context.getDataFormat());
                    LocalDateTime localDateTime;
                    try {
                        localDateTime = LocalDateTime.parse((CharSequence) cellValue, formatter);
                    }catch (DateTimeParseException timeParseException){
                        if (timeParseException.getMessage().contains("Unable to obtain LocalDateTime from")){
                            localDateTime = LocalDate.parse(cellValue.toString(), DateTimeFormatter.ofPattern(Time.YMD_HORIZONTAL_FORMAT_REGEX)).atStartOfDay();
                        }else {
                            throw throwTimeParseException(context,cellValue,timeParseException);
                        }
                    }
                    return dateClass.cast(localDateTime);
                }else if (dateClass == Date.class){
                    cellValue = checkTrimFeature2NewString(excelPolicies,readerConfig,context,cellValue);
                    SimpleDateFormat format = new SimpleDateFormat(context.getDataFormat());
                    try {
                        Date date = Time.parseData(format, cellValue.toString());
                        return dateClass.cast(date);
                    }catch(Exception parseException){
                        throw throwTimeParseException(context,cellValue,parseException);
                    }
                }else if (dateClass == LocalDate.class){
                    cellValue = checkTrimFeature2NewString(excelPolicies,readerConfig,context,cellValue);
                    DateTimeFormatter formatter = DateTimeFormatter.ofPattern(context.getDataFormat());
                    LocalDate localDate;
                    try {
                        localDate = LocalDate.parse((CharSequence) cellValue, formatter);
                    }catch (DateTimeParseException timeParseException){
                        throw throwTimeParseException(context,cellValue,timeParseException);
                    }
                    return dateClass.cast(localDate);
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
                    }else if(dateClass == LocalDate.class){
                        Instant instant = javaDate.toInstant();
                        ZonedDateTime zonedDateTime = instant.atZone(ZoneId.systemDefault());
                        return dateClass.cast(zonedDateTime.toLocalDate());
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

    /**
     * 日期转换异常处理
     * @param context cast上下文信息
     * @param cellValue 单元格数据
     * @param exception 转换异常
     */
    private AxolotlExcelReadException throwTimeParseException(CastContext<?> context, Object cellValue, Exception exception){
        exception.printStackTrace();
        String message = LoggerHelper.format("请指定正确的日期格式,获取为:[%s],尝试转换为:[%s]", cellValue, context.getDataFormat());
        LoggerHelper.debug(LOGGER,LoggerHelper.format(message+",错误信息: [%s]\n",exception.getMessage()));
        return new AxolotlExcelReadException(context,message);
    }

    /**
     * 对单元格数据进行 TRIM_CELL_VALUE 特性处理
     * @param excelPolicies 需排除的特性
     * @param readerConfig 已配置的特性
     * @param context cast相关的上下文信息
     * @param cellValue 单元格数据
     * @return 处理后的单元格数据
     */
    private Object checkTrimFeature2NewString(Map<ExcelReadPolicy, Object> excelPolicies, ReaderConfig<?> readerConfig, CastContext<?> context, Object cellValue){
        if(context.getDataFormat().contains(" ")){
            //format存在空格，不进行去空格处理
            LoggerHelper.debug(LOGGER,LoggerHelper.format("日期格式化:[%s] 带有空格,不进行 [TRIM_CELL_VALUE] 特性处理",context.getDataFormat()));
            return cellValue;
        }
        if (!excelPolicies.containsKey(ExcelReadPolicy.TRIM_CELL_VALUE)) {
            if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.TRIM_CELL_VALUE)) {
                return Regex.convertSingleLine(cellValue.toString());
            }
        }
        return cellValue;
    }


    @Override
    public boolean support(CellType cellType, Class<NT> clazz) {
        return  (cellType == CellType.STRING ||
                cellType == CellType.NUMERIC ||
                cellType == CellType.BLANK) &&
                (clazz == Date.class ||
                clazz == LocalDateTime.class ||
                clazz == LocalDate.class) ;
    }
}

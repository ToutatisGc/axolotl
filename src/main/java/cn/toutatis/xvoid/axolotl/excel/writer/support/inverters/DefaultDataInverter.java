package cn.toutatis.xvoid.axolotl.excel.writer.support.inverters;

import cn.toutatis.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.toutatis.xvoid.toolkit.constant.Time;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDateTime;
import java.util.Date;

public class DefaultDataInverter implements DataInverter<Object> {

    @Override
    public Object convert(Object value) {
        if (value != null){
            Class<?> valueClass = value.getClass();
            if (valueClass == Double.class || valueClass == Float.class){
                return String.format("%.2f",value);
            }
            if (valueClass == BigDecimal.class){
                BigDecimal decimal = (BigDecimal) value;
                boolean isInteger = decimal.remainder(BigDecimal.ONE).compareTo(BigDecimal.ZERO) == 0;
                if (isInteger){
                    return decimal.toBigInteger().toString();
                }else {
                    decimal = decimal.setScale(AxolotlDefaultReaderConfig.XVOID_DEFAULT_DECIMAL_SCALE, RoundingMode.HALF_UP);
                    return decimal.toString();
                }
            }
            if (valueClass == LocalDateTime.class){
                return Time.regexTime(Time.YMD_HORIZONTAL_FORMAT_REGEX,((LocalDateTime) value));
            }
            if (valueClass == Date.class){
                return Time.regexTime(Time.YMD_HORIZONTAL_FORMAT_REGEX,((Date) value));
            }
            return value;
        }
        return null;
    }
}

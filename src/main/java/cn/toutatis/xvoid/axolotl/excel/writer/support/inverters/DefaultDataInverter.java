package cn.toutatis.xvoid.axolotl.excel.writer.support.inverters;

import cn.toutatis.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.support.DataInverter;

import java.math.BigDecimal;
import java.math.RoundingMode;

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
            return value;
        }
        return null;
    }
}

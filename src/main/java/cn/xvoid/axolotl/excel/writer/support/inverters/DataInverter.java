package cn.xvoid.axolotl.excel.writer.support.inverters;

/**
 * 实体类写入应当支持类型变换
 * 目前支持写入字符串
 * @author Toutatis_Gc
 */
public interface DataInverter<T> {

    /**
     * 类型转换
     * @param value 待转换值
     * @return 转换后的值
     */
    T convert(Object value);

}

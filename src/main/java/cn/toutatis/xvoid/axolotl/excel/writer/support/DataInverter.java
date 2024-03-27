package cn.toutatis.xvoid.axolotl.excel.writer.support;

/**
 * TODO
 * 实体类写入应当支持类型变换
 * 目前支持写入字符串
 * @author Toutatis_Gc
 */
public interface DataInverter<T> {

    T convert(Object value);

}

package cn.toutatis.xvoid.axolotl.excel.reader.constant;

/**
 * 该模块使用的默认配置
 * @author Toutatis_Gc
 */
public class AxolotlDefaultReaderConfig {

    /**
     * xlsx默认最大文件大小5M
     * file.length获取文件大小时，获取单位为bytes，这里将其转换为5M
     */
    public static final long XVOID_DEFAULT_MAX_FILE_SIZE = 5*1024*1024;

    /**
     * 读取每批次的行数
     */
    public static final int XVOID_DEFAULT_READ_EACH_BATCH_SIZE = 200;

    /**
     * 默认的小数位数
     */
    public static final int XVOID_DEFAULT_DECIMAL_SCALE = 2;




}

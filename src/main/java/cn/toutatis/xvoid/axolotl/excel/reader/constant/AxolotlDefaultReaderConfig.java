package cn.toutatis.xvoid.axolotl.excel.reader.constant;

/**
 * 该模块使用的默认配置
 * @author Toutatis_Gc
 */
public class AxolotlDefaultReaderConfig {

    /**
     * 读取每批次的行数
     */
    public static final int XVOID_DEFAULT_READ_EACH_BATCH_SIZE = 200;

    /**
     * 默认的小数位数
     */
    public static final int XVOID_DEFAULT_DECIMAL_SCALE = 2;

    /**
     * 默认的表头查找行数
     */
    public static final int XVOID_DEFAULT_HEADER_FINDING_ROW = 10;

    /**
     * 对于流进行读取时，由于无法获取文件本身内容，需要读取Excel文件所对应的压缩内容
     * Excel标志文件为[Content_Types].xml
     */
    public static final String EXCEL_ZIP_XML_FILE_NAME = "[Content_Types].xml";

}

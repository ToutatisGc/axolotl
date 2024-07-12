package cn.xvoid.axolotl.common;

import org.apache.tika.mime.MimeType;
import org.apache.tika.mime.MimeTypeException;
import org.apache.tika.mime.MimeTypes;

/**
 * 常见的MIME类型
 * @author Toutatis_Gc
 */
public class CommonMimeType {

    /**
     * MS_EXCEL Excel 97-2003文件版本
     * application/vnd.ms-excel
     * 后缀为[.xls]
     */
    public static final MimeType MS_EXCEL;

    /**
     * 该类型为读取流的情况获取到的MS_EXCEL格式
     * 使用tika的API获取到的类型为application/x-tika-msoffice
     */
    public static final MimeType TIKA_MS_EXCEL;

    /**
     * OOXML_EXCEL Excel 2007及以上版本
     * application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
     * 后缀为[.xlsx]
     */
    public static final MimeType OOXML_EXCEL;
    public static final MimeType TIKA_OOXML_EXCEL;

    /**
     * ZIP 压缩文件
     * 进行流读取时，Excel文件也为zip格式
     * application/zip
     * 后缀为[.zip]
     */
    public static final MimeType ZIP;

    /**
     * 二进制流
     * application/octet-stream
     */
    public static final MimeType OCTET_STREAM;

    static {
        try {
            MimeTypes defaultMimeTypes = MimeTypes.getDefaultMimeTypes();
            MS_EXCEL = defaultMimeTypes.forName("application/vnd.ms-excel");
            TIKA_MS_EXCEL = defaultMimeTypes.forName("application/x-tika-msoffice");
            OOXML_EXCEL = defaultMimeTypes.forName("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            TIKA_OOXML_EXCEL = defaultMimeTypes.forName("application/x-tika-ooxml");
            ZIP = defaultMimeTypes.forName("application/zip");
            OCTET_STREAM = defaultMimeTypes.forName("application/octet-stream");
        } catch (MimeTypeException e) {
            throw new RuntimeException(e);
        }
    }

}

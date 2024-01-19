package cn.toutatis.xvoid.axolotl.common;

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
     * OOXML_EXCEL Excel 2007及以上版本
     * application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
     * 后缀为[.xlsx]
     */
    public static final MimeType OOXML_EXCEL;

    /**
     * ZIP 压缩文件
     * 进行流读取时，Excel文件也为zip格式
     * application/zip
     * 后缀为[.zip]
     */
    public static final MimeType ZIP;

    static {
        try {
            MimeTypes defaultMimeTypes = MimeTypes.getDefaultMimeTypes();
            MS_EXCEL = defaultMimeTypes.forName("application/vnd.ms-excel");
            OOXML_EXCEL = defaultMimeTypes.forName("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            ZIP = defaultMimeTypes.forName("application/zip");
        } catch (MimeTypeException e) {
            throw new RuntimeException(e);
        }
    }

}

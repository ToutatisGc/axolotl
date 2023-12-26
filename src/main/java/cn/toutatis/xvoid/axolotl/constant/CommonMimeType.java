package cn.toutatis.xvoid.axolotl.constant;

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


    static {
        try {
            MimeTypes defaultMimeTypes = MimeTypes.getDefaultMimeTypes();
            MS_EXCEL = defaultMimeTypes.forName("application/vnd.ms-excel");
            OOXML_EXCEL = defaultMimeTypes.forName("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        } catch (MimeTypeException e) {
            throw new RuntimeException(e);
        }
    }

}

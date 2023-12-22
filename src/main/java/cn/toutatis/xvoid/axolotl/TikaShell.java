package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.support.DetectResult;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import org.apache.tika.Tika;
import org.apache.tika.mime.MimeType;
import org.apache.tika.mime.MimeTypeException;
import org.apache.tika.mime.MimeTypes;
import org.slf4j.Logger;

import java.io.File;
import java.io.IOException;

/**
 * Tika 工具壳
 * @author Toutatis_Gc
 */
public class TikaShell {

    private static final Logger LOGGER = LoggerToolkit.getLogger(TikaShell.class);

    private static final Tika tika = new Tika();

    /**
     * MS_EXCEL Excel 97-2003文件版本
     * application/vnd.ms-excel
     * 后缀为[.xls]
     */
    public static final MimeType MS_EXCEL;

    static {
        try {
            MS_EXCEL = MimeTypes.getDefaultMimeTypes().forName("application/vnd.ms-excel");
        } catch (MimeTypeException e) {
            throw new RuntimeException(e);
        }
    }

    public static DetectResult detect(File file, MimeType mimeType) {
        if (file == null){return new DetectResult(false, "文件不存在");}
        if (FileToolkit.exists(file) && file.isFile()){
            try {
                String detect = tika.detect(file);
                return new DetectResult(true,"TODO" );
            } catch (IOException e) {
                e.printStackTrace();
                String msg = "文件读取失败";
                LOGGER.error(file.getName()+msg, e);
                return new DetectResult(false, msg);
            }
        }else{
            return new DetectResult(false, "文件不存在");
        }

    }

    public static void main(String[] args) {

    }

}

package cn.toutatis.xvoid.axolotl.support;

import cn.toutatis.xvoid.axolotl.constant.CommonMimeType;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import org.apache.tika.Tika;
import org.apache.tika.mime.MimeType;
import org.apache.tika.mime.MimeTypeException;
import org.apache.tika.mime.MimeTypes;
import org.slf4j.Logger;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * Tika 工具壳
 * @author Toutatis_Gc
 */
public class TikaShell {

    private static final Logger LOGGER = LoggerToolkit.getLogger(TikaShell.class);

    /**
     * Tika 实例
     */
    private static final Tika tika = new Tika();

    public static final MimeType MS_EXCEL = CommonMimeType.MS_EXCEL;

    public static final MimeType OOXML_EXCEL = CommonMimeType.OOXML_EXCEL;


    /**
     * 判断文件类型是否为需要的类型
     * @param file 文件
     * @param mimeType 想要匹配的MIME类型
     * @return 检测结果
     */
    public static DetectResult detect(File file, MimeType mimeType) {
        DetectResult preCheck = preCheck(file);
        if (preCheck.isDetect()){
            try {
                String fileSuffix = '.'+FileToolkit.getFileSuffix(file).toLowerCase();
                List<String> extensions = mimeType.getExtensions();
                int idx = -1;
                LOGGER.debug("文件后缀：["+ fileSuffix + "], 可匹配的后缀：" + mimeType.getExtensions().toString() + ", 匹配的后缀索引：[" + idx + "]");
                for (int i = 0; i < extensions.size(); i++) {
                    String extension = extensions.get(i);
                    if (extension.equalsIgnoreCase(fileSuffix)){
                        idx = i;
                        break;
                    }
                }
                if (idx == -1){
                    return new DetectResult(false, "文件后缀不匹配");
                }
                if (tika.detect(file).equals(mimeType.toString())){
                    return new DetectResult(true);
                }else {
                    return new DetectResult(false, "类型不匹配");
                }
            } catch (IOException e) {
                e.printStackTrace();
                String msg = "文件读取失败";
                LOGGER.error(file.getName()+msg, e);
                return new DetectResult(false, msg);
            }
        }else{
            return preCheck;
        }
    }

    /**
     * 获取文件类型
     * @param file 文件
     * @return MIME类型
     */
    public static MimeType getMimeType(File file){
        DetectResult preCheck = preCheck(file);
        if (preCheck.isDetect()){
            try {
                return MimeTypes.getDefaultMimeTypes().forName(tika.detect(file));
            } catch (IOException | MimeTypeException e) {
                e.printStackTrace();
                String msg = "文件读取失败";
                LOGGER.error(file.getName()+msg, e);
                throw new RuntimeException(msg);
            }
        }else{
            String msg = "文件读取失败";
            LOGGER.error(file.getName()+preCheck.getMessage());
            throw new RuntimeException(preCheck.getMessage());
        }
    }

    /**
     * 预检查
     * @param file 文件
     * @return 检测结果
     */
    private static DetectResult preCheck(File file){
        if (file == null || !file.exists()){
            return new DetectResult(false, "文件不存在");
        }
        if (file.isDirectory()){
            return new DetectResult(false, "选择文件不能是目录");
        }else{
            return new DetectResult(true);
        }
    }

}

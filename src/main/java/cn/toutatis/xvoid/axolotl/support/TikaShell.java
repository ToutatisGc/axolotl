package cn.toutatis.xvoid.axolotl.support;

import cn.toutatis.xvoid.axolotl.constant.CommonMimeType;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import lombok.SneakyThrows;
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


    public static DetectResult detect(File file, MimeType mimeType){
        return detect(file, mimeType, false);
    }
    public static DetectResult detectThrowException(File file, MimeType mimeType){
        return detect(file, mimeType, true);
    }
    /**
     * 判断文件类型是否为需要的类型
     * @param file 文件
     * @param mimeType 想要匹配的MIME类型
     * @return 检测结果
     */
    @SneakyThrows
    public static DetectResult detect(File file, MimeType mimeType, boolean throwException) {
        DetectResult preCheck = preCheckFileNormal(file);
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
                    String msg = file.getName()+"文件后缀不匹配";
                    if (throwException){throw new IOException(msg);}
                    return new DetectResult(false, DetectResult.FileStatus.FILE_META_PROBLEM, msg);
                }
                if (tika.detect(file).equals(mimeType.toString())){
                    return new DetectResult(true);
                }else {
                    String msg = (file.getName()+"文件媒体类型不匹配，媒体类型：" + tika.detect(file) + ", 期望媒体类型：" + mimeType.toString());
                    if (throwException){throw new IOException(msg);}
                    return new DetectResult(false, DetectResult.FileStatus.FILE_META_PROBLEM, msg);
                }
            } catch (IOException e) {
                e.printStackTrace();
                String msg = file.getName()+"文件读取失败";
                if (throwException){throw new IOException(msg);}
                LOGGER.error(msg, e);
                return new DetectResult(false, DetectResult.FileStatus.FILE_META_PROBLEM, msg);
            }
        }else{
            if (throwException){throw new IOException(preCheck.getMessage());}
            return preCheck;
        }
    }

    /**
     * 获取文件类型
     * @param file 文件
     * @return MIME类型
     */
    @SneakyThrows
    public static MimeType getMimeType(File file){
        DetectResult preCheck = preCheckFileNormal(file);
        if (preCheck.isDetect()){
            try {
                return MimeTypes.getDefaultMimeTypes().forName(tika.detect(file));
            } catch (IOException | MimeTypeException e) {
                e.printStackTrace();
                String msg = "文件读取失败";
                LOGGER.error(file.getName()+msg, e);
                throw new IOException(msg);
            }
        }else{
            LOGGER.error(file.getName()+preCheck.getMessage());
            throw new IOException(preCheck.getMessage());
        }
    }

    /**
     * 预检查文件是否正常
     * @param file 文件
     * @return 检测结果
     */
    private static DetectResult preCheckFileNormal(File file){
        if (file == null || !file.exists()){
            return new DetectResult(false, DetectResult.FileStatus.FILE_SELF_PROBLEM, "文件不存在");
        }
        if (file.isDirectory()){
            return new DetectResult(false, DetectResult.FileStatus.FILE_SELF_PROBLEM,"选择文件不能是目录");
        }else{
            return new DetectResult(true);
        }
    }

}

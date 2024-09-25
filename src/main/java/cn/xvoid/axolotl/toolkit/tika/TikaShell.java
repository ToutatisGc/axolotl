package cn.xvoid.axolotl.toolkit.tika;

import cn.xvoid.axolotl.common.CommonMimeType;
import cn.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.xvoid.axolotl.toolkit.LoggerHelper;
import cn.xvoid.toolkit.file.FileToolkit;
import cn.xvoid.toolkit.log.LoggerToolkit;
import com.google.common.io.ByteStreams;
import lombok.SneakyThrows;
import org.apache.tika.Tika;
import org.apache.tika.mime.MimeType;
import org.apache.tika.mime.MimeTypeException;
import org.apache.tika.mime.MimeTypes;
import org.slf4j.Logger;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * Tika 工具壳
 * @author Toutatis_Gc
 */
public class TikaShell {

    /**
     * 日志工具
     */
    private static final Logger LOGGER = LoggerToolkit.getLogger(TikaShell.class);

    /**
     * Tika 实例
     */
    private static final Tika tika = new Tika();

    /**
     * MS_EXCEL Excel 97-2003文件版本
     * application/vnd.ms-excel
     * 后缀为[.xls]
     */
    public static final MimeType MS_EXCEL = CommonMimeType.MS_EXCEL;

    /**
     * OOXML_EXCEL Excel 2007及以上版本
     * application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
     * 后缀为[.xlsx]
     */
    public static final MimeType OOXML_EXCEL = CommonMimeType.OOXML_EXCEL;


    /**
     * 检测给定文件是否符合指定的MIME类型
     * 此方法重载了detect方法，简化调用时不需要预先检查的场景
     *
     * @param file 要检测的文件
     * @param mimeType 指定的MIME类型
     * @return 检测结果
     */
    public static DetectResult detect(File file, MimeType mimeType){
        return detect(file, mimeType, false);
    }

    /**
     * 检测给定文件是否符合指定的MIME类型，允许设置是否已预先检查
     * 此方法提供了预先检查的选项，以便在已知进行了预检查的情况下避免重复检查
     *
     * @param file 要检测的文件
     * @param mimeType 指定的MIME类型
     * @param alreadyPreCheck 是否已经预先检查
     * @return 检测结果
     */
    public static DetectResult detect(File file, MimeType mimeType,boolean alreadyPreCheck){
        return detect(file, mimeType, false,alreadyPreCheck);
    }

    /**
     * 检测给定文件是否符合指定的MIME类型，并在检测失败时抛出异常
     * 此方法适用于对文件有严格类型要求的场景，通过抛出异常来处理不合规的文件
     *
     * @param file 要检测的文件
     * @param mimeType 指定的MIME类型
     * @return 检测结果
     * @throws IllegalArgumentException 如果文件不符合指定的MIME类型
     */
    public static DetectResult detectThrowException(File file, MimeType mimeType){
        return detect(file, mimeType, true,false);
    }

    /**
     * 判断文件是否正常并且为需要的格式
     * 1.文件预检查
     * 2.文件后缀是否匹配
     * 3.文件媒体类型是否匹配
     *
     * @param file 文件
     * @param wantedMimeType 想要匹配的MIME类型
     * @param throwException 是否抛出异常
     * @param alreadyPreCheck 文件是否已通过预检查
     * @return 检测结果
     */
    @SneakyThrows
    public static DetectResult detect(File file, MimeType wantedMimeType, boolean throwException, boolean alreadyPreCheck) {
        DetectResult preCheck;
        if (alreadyPreCheck){
            //已通过预检测
            preCheck = new DetectResult(true);
        }else{
            //进行预检测
            preCheck = preCheckFileNormal(file);
        }
        if (preCheck.isDetect()){
            //进行文件后缀匹配
            DetectResult detectResult = new DetectResult(false);
            detectResult.setWantedMimeType(wantedMimeType);
            //获取文件后缀
            String fileSuffix = '.'+FileToolkit.getFileSuffix(file).toLowerCase();
            //获取期望的文件后缀
            List<String> extensions = wantedMimeType.getExtensions();
            //进行后缀匹配
            int idx = -1;
            for (int i = 0; i < extensions.size(); i++) {
                String extension = extensions.get(i);
                if (extension.equalsIgnoreCase(fileSuffix)){
                    idx = i;
                    break;
                }
            }
            LOGGER.debug("文件后缀：["+ fileSuffix + "], 可匹配的后缀：" + wantedMimeType.getExtensions() + ", 匹配的后缀索引：[" + idx + "]");
            if (idx == -1){
                //文件后缀不匹配 返回错误信息
                String msg ="["+file.getName()+"]文件后缀不匹配";
                detectResult.setCurrentFileStatus(DetectResult.FileStatus.FILE_SUFFIX_PROBLEM);
                if (throwException){throw new IOException(msg);}
                return detectResult.returnInfo(msg);
            }
            String detect = tika.detect(file);
            return matchMimeType(detect, wantedMimeType,throwException);
        }else{
            //再次预检测失败 返回错误信息
            if (throwException){throw new IOException(preCheck.getMessage());}
            return preCheck;
        }
    }

    /**
     * 判断文件是否正常并且为需要的格式
     * 流的情况较为特殊，由于流获取内容仅为字节，在获取文件类型时仅能获取较为特殊的几种类型，因此需要进行额外的判断
     * Excel文件本身为ZIP压缩类型，其中有特殊的标志文件可以进行判断，其余特殊类型请自行实现判断或提交PR进行类型补充
     * @param ins 文件流
     * @param wantedMimeType 想要匹配的MIME类型
     * @param throwException 是否抛出异常
     * @return 检测结果
     */
    @SneakyThrows
    public static DetectResult detect(InputStream ins, MimeType wantedMimeType, boolean throwException) {
        if (wantedMimeType == null){throw new IllegalArgumentException("期望媒体类型为空");}
        ByteArrayOutputStream dataCacheOutputStream =  new ByteArrayOutputStream();
        ByteStreams.copy(ins, dataCacheOutputStream);
        //解析出文件的媒体类型
        String detectMimeTypeString = tika.detect(new ByteArrayInputStream(dataCacheOutputStream.toByteArray()));
        if (wantedMimeType == CommonMimeType.MS_EXCEL || wantedMimeType == CommonMimeType.OOXML_EXCEL){
            if (CommonMimeType.ZIP.toString().equals(detectMimeTypeString)){
                ZipEntry entry;
                try (ZipInputStream zipInputStream = new ZipInputStream(new ByteArrayInputStream(dataCacheOutputStream.toByteArray()))){
                    while ((entry = zipInputStream.getNextEntry()) != null) {
                        if (entry.isDirectory()){continue;}
                        if (entry.getName().equals(AxolotlDefaultReaderConfig.EXCEL_ZIP_XML_FILE_NAME)){
                            ByteArrayOutputStream tmpOutputStream = new ByteArrayOutputStream();
                            ByteStreams.copy(zipInputStream, tmpOutputStream);
                            String fileContent = new String(tmpOutputStream.toByteArray(),StandardCharsets.UTF_8);
                            if (fileContent.contains(wantedMimeType.toString())){
                                detectMimeTypeString = wantedMimeType.toString();
                            }
                        }
                    }
                }
            } else if (CommonMimeType.TIKA_OOXML_EXCEL.toString().equals(detectMimeTypeString)) {
                detectMimeTypeString = CommonMimeType.OOXML_EXCEL.toString();
            } else if (CommonMimeType.TIKA_MS_EXCEL.toString().equals(detectMimeTypeString)) {
                detectMimeTypeString = CommonMimeType.MS_EXCEL.toString();
            } else if (CommonMimeType.OCTET_STREAM.toString().equals(detectMimeTypeString)){
                // 二进制流的情况复杂，无法判断类型，只能是将文件尝试读取
                detectMimeTypeString = wantedMimeType.toString();
            }else {
                String msg = "流不是Excel文件";
                if (throwException){throw new IOException(msg);}
                DetectResult detectResult = new DetectResult(false);
                detectResult.setWantedMimeType(wantedMimeType);
                detectResult.returnInfo(msg);
            }
        }
        try {
            return matchMimeType(detectMimeTypeString, wantedMimeType, throwException);
        } catch (IOException e) {
            String msg = "文件读取失败";
            if (throwException){throw new IOException(msg);}
            LOGGER.error(msg, e);
            DetectResult detectResult = new DetectResult(false, DetectResult.FileStatus.FILE_SELF_PROBLEM, msg);
            detectResult.setWantedMimeType(wantedMimeType);
            return detectResult;
        }
    }

    /**
     * 匹配文件媒体类型
     * @param detectMimeTypeString 文件媒体类型
     * @param wantedMimeType 期望媒体类型
     * @param throwException 是否抛出异常
     * @return 识别结果
     * @throws MimeTypeException MIME类型异常
     * @throws IOException 文件读取失败
     */
    private static DetectResult matchMimeType(String detectMimeTypeString, MimeType wantedMimeType, boolean throwException) throws MimeTypeException, IOException {
        MimeType detectMimeType = MimeTypes.getDefaultMimeTypes().forName(detectMimeTypeString);
        LoggerHelper.debug(
                LOGGER, LoggerHelper.format("文件媒体类型：%s 期望媒体类型：%s" ,detectMimeTypeString, wantedMimeType)
        );
        DetectResult detectResult = new DetectResult(false);
        if (detectMimeType.equals(wantedMimeType)){
            //文件媒体类型与期望媒体类型一致 返回检测结果
            detectResult = new DetectResult(true,detectMimeType);
        }else {
            //不一致  返回错误信息
            detectResult.setCatchMimeType(detectMimeType);
            detectResult.setCurrentFileStatus(DetectResult.FileStatus.FILE_MIME_TYPE_PROBLEM);
            String msg = ("文件媒体类型不匹配，媒体类型：" + detectMimeTypeString + ", 期望媒体类型：" + wantedMimeType);
            if (throwException){throw new IOException(msg);}
            return detectResult.returnInfo(msg);
        }
        return detectResult;
    }

    /**
     * 获取文件类型
     *
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
     * 预检查文件是否正常并抛出异常
     *
     * @param file 文件
     * @return 检测结果
     */
    public static DetectResult preCheckFileNormalThrowException(File file){
        return preCheckFileNormal(file,true);
    }

    /**
     * 预检查文件是否正常
     *
     * @param file 文件
     * @return 检测结果
     */
    public static DetectResult preCheckFileNormal(File file){
        return preCheckFileNormal(file,false);
    }

    /**
     * 预检查文件是否正常
     *
     * @param file 文件
     * @param throwException 是否抛出异常
     * @return 检测结果
     */
    @SneakyThrows
    public static DetectResult preCheckFileNormal(File file,boolean throwException){
        if (file == null || !file.exists()){
            String msg = "文件不存在";
            if (throwException){throw new FileNotFoundException(msg);}
            return new DetectResult(false, DetectResult.FileStatus.FILE_SELF_PROBLEM, msg);
        }
        if (file.isDirectory()){
            String msg = "选择文件不能是目录";
            if (throwException){throw new IOException(msg);}
            return new DetectResult(false, DetectResult.FileStatus.FILE_SELF_PROBLEM,msg);
        }else{
            return new DetectResult(true);
        }
    }

}

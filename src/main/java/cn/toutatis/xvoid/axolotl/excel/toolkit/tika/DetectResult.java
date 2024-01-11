package cn.toutatis.xvoid.axolotl.excel.toolkit.tika;


import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.Getter;
import lombok.SneakyThrows;
import org.apache.tika.mime.MimeType;

import java.io.IOException;
import java.util.Objects;

/**
 * 文件检测结果
 */
@Data
@AllArgsConstructor
public class DetectResult {

    /**
     * 是否识别通过
     */
    private Boolean detect;

    /**
     * 文件检测结果
     */
    private String message;

    /**
     * 识别到的文件类型
     */
    private MimeType catchMimeType;

    /**
     * 期望的文件类型
     */
    private MimeType wantedMimeType;

    /**
     * 文件读取状态
     */
    private FileStatus currentFileStatus;

    /**
     * 返回文件状态信息
     * @param msg 状态信息
     * @return 检测对象
     */
    public DetectResult returnInfo(String msg) {
        this.message = msg;
        return this;
    }

    @Getter
    @AllArgsConstructor
    public enum FileStatus {
        /**
         * 识别通过
         */
        DETECTED(0,"文件正常"),
        /**
         * 文件自身问题识别不通过
         */
        FILE_SELF_PROBLEM(1,"文件自身问题"),
        /**
         * 文件内容格式识别不通过
         */
        FILE_MIME_TYPE_PROBLEM(2,"文件媒体类型问题"),
        /**
         * 文件后缀识别不通过
         */
        FILE_SUFFIX_PROBLEM(3,"文件后缀问题");

        /**
         * 状态码
         */
        private final int idx;

        /**
         * 状态信息
         */
        private final String msg;

    }

    public DetectResult(boolean detect) {
        this.detect = detect;
        this.currentFileStatus = FileStatus.DETECTED;
    }

    public DetectResult(boolean detect,MimeType wantedMimeType) {
        this.detect = detect;
        this.currentFileStatus = FileStatus.DETECTED;
        this.setWantedMimeType(wantedMimeType);
        if (detect){
           this.setCatchMimeType(wantedMimeType);
        }
    }

    public DetectResult(boolean detect, FileStatus currentFileStatus, String message) {
        this.detect = detect;
        this.currentFileStatus = currentFileStatus;
        this.message = message;
    }

    /**
     * 文件自身是否识别有问题
     * @return 文件不存在或者为目录等问题
     */
    public boolean isFileSelfProblem() {
        return !getDetect() && currentFileStatus == FileStatus.FILE_SELF_PROBLEM;
    }

    /**
     * 文件格式是否识别有问题
     * @return 文件格式识别有问题
     */
    public boolean isFileMetaProblem() {
        return!getDetect() && currentFileStatus == FileStatus.FILE_MIME_TYPE_PROBLEM;
    }

    /**
     * 文件是否识别成功
     * @return 是否识别成功
     */
    public boolean isDetect() {
        if (Objects.isNull(detect)){
            return isWantedMimeType();
        }else return detect;
    }

    /**
     * 文件媒体类型是所需类型
     * @return 文件媒体类型是所需类型
     */
    public boolean isWantedMimeType() {
        return wantedMimeType.equals(catchMimeType);
    }

    @SneakyThrows
    public void throwException(){
        throw new IOException(message);
    }
}

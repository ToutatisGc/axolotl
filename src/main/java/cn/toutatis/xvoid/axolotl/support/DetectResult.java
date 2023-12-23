package cn.toutatis.xvoid.axolotl.support;


import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.tika.mime.MimeType;

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

    public enum FileStatus {
        /**
         * 识别通过
         */
        DETECTED,
        /**
         * 文件自身问题识别不通过
         */
        FILE_SELF_PROBLEM,
        /**
         * 文件内容问题识别不通过
         */
        FILE_META_PROBLEM
    }

    public DetectResult(boolean detect) {
        this.detect = detect;
        this.currentFileStatus = FileStatus.DETECTED;
    }

    public DetectResult(boolean detect, FileStatus currentFileStatus, String message) {
        this.detect = detect;
        this.currentFileStatus = currentFileStatus;
        this.message = message;
    }

    public boolean isFileSelfProblem() {
        return !getDetect() && currentFileStatus == FileStatus.FILE_SELF_PROBLEM;
    }

    public boolean isDetect() {
        return Objects.requireNonNullElseGet(detect, this::isWantedMimeType);
    }

    public boolean isWantedMimeType() {
        return catchMimeType.equals(wantedMimeType);
    }
}

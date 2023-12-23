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

    public DetectResult(boolean detect) {this.detect = detect;}

    public DetectResult(boolean detect, String message) {
        this.detect = detect;
        this.message = message;
    }

    public Boolean isDetect() {
        return Objects.requireNonNullElseGet(detect, () -> wantedMimeType.equals(catchMimeType));
    }

    public Boolean isWantedMimeType() {
        return catchMimeType.equals(wantedMimeType);
    }
}

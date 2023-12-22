package cn.toutatis.xvoid.axolotl.support;


import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * 文件检测结果
 */
@Data
@AllArgsConstructor
public class DetectResult {

    /**
     * 是否识别通过
     */
    private boolean detect = false;

    /**
     * 文件检测结果
     */
    private String message;

    public DetectResult(boolean detect) {this.detect = detect;}
}

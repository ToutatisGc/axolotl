package cn.toutatis.xvoid.axolotl.excel.writer.support;

import lombok.Data;

/**
 * 写入结果
 * @author Toutatis_Gc
 */
@Data
public class AxolotlWriteResult {
    public AxolotlWriteResult() {
    }

    public AxolotlWriteResult(boolean write, String message) {
        this.write = write;
        this.message = message;
    }

    /**
     * 是否写入成功
     */
    private boolean write = false;

    /**
     * 写入结果信息
     */
    private String message;

}

package cn.toutatis.xvoid.axolotl.excel.writer.support;

import lombok.Data;

/**
 * 写入结果
 * @author Toutatis_Gc
 */
@Data
public class AxolotlWriteResult {

    /**
     * 是否写入成功
     */
    private boolean write = false;

    /**
     * 写入结果信息
     */
    private String message;

}

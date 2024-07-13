package cn.xvoid.axolotl.excel.writer.exceptions;

import cn.xvoid.axolotl.exceptions.AxolotlException;

/**
 * 写入时的异常
 */
public class AxolotlWriteException extends AxolotlException {

    public AxolotlWriteException(String message) {
        super(message);
    }

    public AxolotlWriteException(Throwable cause) {
        super(cause);
    }
}

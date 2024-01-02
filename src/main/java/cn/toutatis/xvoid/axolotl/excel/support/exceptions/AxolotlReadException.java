package cn.toutatis.xvoid.axolotl.excel.support.exceptions;

/**
 * 读取Excel异常统一抛出
 */
public class AxolotlReadException extends RuntimeException{

    public AxolotlReadException(Throwable cause) {
        super(cause);
    }

    public AxolotlReadException(String message) {
        super(message);
    }
}

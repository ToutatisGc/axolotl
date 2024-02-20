package cn.toutatis.xvoid.axolotl.exceptions;

/**
 * 框架一般错误
 * @author Toutatis_Gc
 */
public class AxolotlException extends RuntimeException {

    public AxolotlException(String message) {
        super(message);
    }

    public AxolotlException(Throwable cause) {
        super(cause);
    }
}

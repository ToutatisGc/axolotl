package cn.toutatis.xvoid.axolotl.excel.support.exceptions;

/**
 * 读取Excel异常统一抛出
 * @author Toutatis_Gc
 */
public class AxolotlExcelReadException extends RuntimeException{

    public AxolotlExcelReadException(Throwable cause) {
        super(cause);
    }

    public AxolotlExcelReadException(String message) {
        super(message);
    }
}

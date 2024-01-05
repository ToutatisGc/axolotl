package cn.toutatis.xvoid.axolotl.excel.support.exceptions;

import cn.toutatis.xvoid.axolotl.excel.WorkBookContext;
import cn.toutatis.xvoid.axolotl.excel.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.toolkit.ExcelToolkit;
import lombok.Getter;
import lombok.Setter;

/**
 * 读取Excel异常统一抛出
 * @author Toutatis_Gc
 */
@Getter @Setter
public class AxolotlExcelReadException extends RuntimeException{

    /**
     * 当前读取的行号
     */
    private int currentReadRowIndex;

    /**
     * 当前读取的列号
     */
    private int currentReadColumnIndex;

    public AxolotlExcelReadException(Throwable cause) {
        super(cause);
    }

    public AxolotlExcelReadException(String message) {
        super(message);
    }

    public AxolotlExcelReadException(WorkBookContext workBookContext, String message) {
        this(workBookContext.getCurrentReadRowIndex(), workBookContext.getCurrentReadColumnIndex(), message);
    }

    public AxolotlExcelReadException(CastContext<?> castContext, String message) {
        this(castContext.getCurrentReadRowIndex(), castContext.getCurrentReadColumnIndex(), message);
    }

    public AxolotlExcelReadException(int currentReadRowIndex, int currentReadColumnIndex, String message){
        super(message);
        this.setCurrentReadColumnIndex(currentReadColumnIndex);
        this.setCurrentReadRowIndex(currentReadRowIndex);
    }

    public String getHumanReadablePosition(){
        return ExcelToolkit.getHumanReadablePosition(currentReadRowIndex, currentReadColumnIndex);
    }
}

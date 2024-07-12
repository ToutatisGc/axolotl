package cn.xvoid.axolotl.excel.reader.support.exceptions;

import cn.xvoid.axolotl.excel.reader.WorkBookContext;
import cn.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.xvoid.axolotl.excel.reader.support.CastContext;
import cn.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.xvoid.axolotl.exceptions.AxolotlException;
import lombok.Getter;
import lombok.Setter;

/**
 * 读取Excel异常统一抛出
 * @author Toutatis_Gc
 */
@Getter @Setter
public class AxolotlExcelReadException extends AxolotlException {

    /**
     * 当前读取的行号
     */
    private int currentReadRowIndex;

    /**
     * 当前读取的列号
     */
    private int currentReadColumnIndex;

    /**
     * 错误的属性名称
     */
    private String fieldName;

    /**
     * 异常类型
     */
    private ExceptionType exceptionType;

    public enum ExceptionType{

        /**
         * 读取Excel文件时出现了异常
         */
        READ_EXCEL_ERROR,

        /**
         * 读取Excel数据时，出现了异常
         */
        READ_EXCEL_ROW_ERROR,

        /**
         * 转换数据时出现异常
         */
        CONVERT_FIELD_ERROR,

        /**
         * 校验数据时出现异常
         */
        VALIDATION_ERROR
    }

    public AxolotlExcelReadException(ExceptionType exceptionType,Throwable cause) {
        super(cause);
    }

    public AxolotlExcelReadException(ExceptionType exceptionType,String message) {
        super(message);
    }

    public AxolotlExcelReadException(WorkBookContext workBookContext, String message) {
        this(
                ExceptionType.READ_EXCEL_ROW_ERROR,null,
                workBookContext.getCurrentReadRowIndex(), workBookContext.getCurrentReadColumnIndex(), message
        );
    }

    public AxolotlExcelReadException(EntityCellMappingInfo<?> mappingInfo, String message){
        this(
                ExceptionType.READ_EXCEL_ROW_ERROR, mappingInfo.getFieldName(),
                mappingInfo.getRowPosition(), mappingInfo.getColumnPosition(), message
        );
    }

    public AxolotlExcelReadException(CastContext<?> castContext, String message) {
        this(
                ExceptionType.CONVERT_FIELD_ERROR,castContext.getCastType().getSimpleName(),
                castContext.getCurrentReadRowIndex(), castContext.getCurrentReadColumnIndex(), message
        );
    }

    public AxolotlExcelReadException(
            ExceptionType exceptionType,String fieldName,
            int currentReadRowIndex, int currentReadColumnIndex, String message
    ){
        super(message);
        this.exceptionType = exceptionType;
        this.fieldName = fieldName;
        this.setCurrentReadColumnIndex(currentReadColumnIndex);
        this.setCurrentReadRowIndex(currentReadRowIndex);
    }

    public String getHumanReadablePosition(){
        return ExcelToolkit.getHumanReadablePosition(currentReadRowIndex, currentReadColumnIndex);
    }
}

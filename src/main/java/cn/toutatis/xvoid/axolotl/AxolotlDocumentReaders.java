package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.excel.AxolotlExcelReader;

import java.io.File;

/**
 * 文档加载器
 * @author Toutatis_Gc
 */
public class AxolotlDocumentReaders {

    public static <T> AxolotlExcelReader<T> getExcelReader(File excelFile, Class<T> clazz){
        return new AxolotlExcelReader<>(excelFile, clazz);
    }

    public static AxolotlExcelReader<Object> getExcelReader(File excelFile){
        return getExcelReader(excelFile, Object.class);
    }

}

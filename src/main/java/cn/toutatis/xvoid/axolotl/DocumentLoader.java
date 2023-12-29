package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.excel.GracefulExcelReader;

import java.io.File;

/**
 * 文档加载器
 * @author Toutatis_Gc
 */
public class DocumentLoader {

    public static <T> GracefulExcelReader<T> getExcelReader(File excelFile, Class<T> clazz){
        return new GracefulExcelReader<>(excelFile, clazz);
    }

    public static GracefulExcelReader<Object> getExcelReader(File excelFile){
        return getExcelReader(excelFile, Object.class);
    }

}

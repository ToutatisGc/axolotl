package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlExcelWriter;

import java.io.File;

/**
 * 文档加载器
 * @author Toutatis_Gc
 */
public class Axolotls {

    public static <T> AxolotlExcelReader<T> getExcelReader(File excelFile, Class<T> clazz){
        return new AxolotlExcelReader<>(excelFile, clazz);
    }

    public static AxolotlExcelReader<Object> getExcelReader(File excelFile){
        return new AxolotlExcelReader<>(excelFile);
    }

    public static AxolotlExcelWriter getExcelWriter(File outputFile){
        return new AxolotlExcelWriter(outputFile);
    }

}

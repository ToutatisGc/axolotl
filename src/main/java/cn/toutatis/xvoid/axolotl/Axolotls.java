package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlStreamExcelReader;

import java.io.File;
import java.io.InputStream;

/**
 * 文档加载器静态构造
 * @author Toutatis_Gc
 */
public class Axolotls {

    /**
     * 获取Excel读取器
     * @param excelFile Excel文件
     * @param clazz 数据POJO类
     * @return Excel读取器
     * @param <T> 数据POJO类
     */
    public static <T> AxolotlExcelReader<T> getExcelReader(File excelFile, Class<T> clazz){
        return new AxolotlExcelReader<>(excelFile, clazz);
    }

    /**
     * 获取Excel读取器
     * @param ins Excel文件文件流
     * @param clazz 数据POJO类
     * @return Excel读取器
     * @param <T> 数据POJO类
     */
    public static <T> AxolotlExcelReader<T> getExcelReader(InputStream ins, Class<T> clazz){
        return new AxolotlExcelReader<>(ins, clazz);
    }

    /**
     * 获取无泛型Excel读取器
     * @param excelFile Excel文件
     * @return Excel读取器
     */
    public static AxolotlExcelReader<Object> getExcelReader(File excelFile){
        return new AxolotlExcelReader<>(excelFile);
    }

    /**
     * 获取无泛型Excel读取器
     * @param ins Excel文件文件流
     * @return Excel读取器
     */
    public static AxolotlExcelReader<Object> getExcelReader(InputStream ins){
        return new AxolotlExcelReader<>(ins);
    }

    /**
     * 获取无泛型Excel流读取器
     * @param excelFile Excel文件文件流
     * @return Excel读取器
     */
    public static AxolotlStreamExcelReader<Object> getStreamExcelReader(File excelFile){
        return new AxolotlStreamExcelReader<>(excelFile);
    }
}

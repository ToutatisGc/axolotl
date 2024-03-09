package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlStreamExcelReader;
import cn.toutatis.xvoid.axolotl.excel.writer.*;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

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

    /**
     * 获取模板Excel写入器
     * @param template 模板文件
     * @param writeConfig 写入器配置
     * @return Excel写入器
     */
    public static AxolotlTemplateExcelWriter getTemplateExcelWriter(File template, TemplateWriteConfig writeConfig){
        return new AxolotlTemplateExcelWriter(template, writeConfig);
    }

    /**
     * 获取模板Excel写入器
     * 默认写入第0张表，采用默认策略管理
     * @param template 模板文件
     * @param outputStream 输出流
     * @return Excel写入器
     */
    public static AxolotlTemplateExcelWriter getTemplateExcelWriter(File template, OutputStream outputStream){
        TemplateWriteConfig templateWriteConfig = new TemplateWriteConfig();
        templateWriteConfig.setOutputStream(outputStream);
        return getTemplateExcelWriter(template, templateWriteConfig);
    }

    /**
     * [语法糖]
     * 直接调取模板写入器写入数据
     * @param template 模板文件
     * @param outputStream 输出流
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @return 写入结果
     */
    public static AxolotlWriteResult templateWrite(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas){
        TemplateWriteConfig templateWriteConfig = new TemplateWriteConfig();
        templateWriteConfig.setOutputStream(outputStream);
        AxolotlTemplateExcelWriter templateExcelWriter = getTemplateExcelWriter(template, templateWriteConfig);
        AxolotlWriteResult result = templateExcelWriter.write(fixMapping, datas);
        templateExcelWriter.close();
        return result;
    }

    /**
     * 获取Excel写入器
     * @param writeConfig 写入器配置
     * @return Excel写入器
     */
    public static AxolotlAutoExcelWriter getAutoExcelWriter(AutoWriteConfig writeConfig){
        return new AxolotlAutoExcelWriter(writeConfig);
    }
}

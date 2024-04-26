package cn.toutatis.xvoid.axolotl;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlStreamExcelReader;
import cn.toutatis.xvoid.axolotl.excel.writer.*;
import cn.toutatis.xvoid.axolotl.excel.writer.components.configuration.AxolotlCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.configuration.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.AxolotlSelectBox;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.configurable.AxolotlConfigurableTheme;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.time.LocalDate;
import java.util.ArrayList;
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
     * 获取Excel写入器
     * @param writeConfig 写入器配置
     * @return Excel写入器
     */
    public static AxolotlAutoExcelWriter getAutoExcelWriter(AutoWriteConfig writeConfig){
        return new AxolotlAutoExcelWriter(writeConfig);
    }

}

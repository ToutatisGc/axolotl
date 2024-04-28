package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.excel.writer.support.SheetDataPackage;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlTemplateExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.TemplateWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.util.*;

/**
 * 快拆工具类
 * @author 张智凯
 * @version 1.0.15
 */
public class AxolotlFaster {

    /**
     * [语法糖]
     * 直接调取模板写入器写入数据
     * @param template 模板文件
     * @param outputStream 输出流
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill, boolean templateNonTemplateCellFill){
        TemplateWriteConfig templateWriteConfig = new TemplateWriteConfig();
        templateWriteConfig.setOutputStream(outputStream);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL,templateNullValueWithTemplateFill);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL,templateNonTemplateCellFill);
        AxolotlTemplateExcelWriter templateExcelWriter = Axolotls.getTemplateExcelWriter(template, templateWriteConfig);
        AxolotlWriteResult result = templateExcelWriter.write(fixMapping, datas);
        templateExcelWriter.close();
        return result;
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill){
       return writeToTemplate(template,outputStream,fixMapping,datas,templateNullValueWithTemplateFill,true);
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas){
        return writeToTemplate(template,outputStream,fixMapping,datas,true,true);
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas, boolean templateNullValueWithTemplateFill, boolean templateNonTemplateCellFill){
        return writeToTemplate(template,outputStream,null,datas,templateNullValueWithTemplateFill,templateNonTemplateCellFill);
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas, boolean templateNullValueWithTemplateFill){
        return writeToTemplate(template,outputStream,null,datas,templateNullValueWithTemplateFill,true);
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas){
        return writeToTemplate(template,outputStream,null,datas,true,true);
    }

    /**
     * [语法糖]
     * 直接调取模板写入器写入数据
     * @param template 模板文件
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateWriteConfig 模板写入配置
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template,Map<String,?> fixMapping, List<?> datas, TemplateWriteConfig templateWriteConfig){
        AxolotlTemplateExcelWriter templateExcelWriter = Axolotls.getTemplateExcelWriter(template, templateWriteConfig);
        AxolotlWriteResult result = templateExcelWriter.write(fixMapping, datas);
        templateExcelWriter.close();
        return result;
    }


    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoCatchColumnLength, boolean autoInsertSerialNumber){
        try {
            AutoWriteConfig autoWriteConfig = new AutoWriteConfig();
            autoWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,autoInsertSerialNumber);
            autoWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING,autoInsertTotalInEnding);
            autoWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH,autoCatchColumnLength);
            if(styleRender != null){
                autoWriteConfig.setStyleRender(styleRender);
            }
            if(StringUtils.isNotEmpty(sheetName)){
                autoWriteConfig.setSheetName(sheetName);
            }
            if(StringUtils.isNotEmpty(fontName)){
                autoWriteConfig.setFontName(fontName);
            }
            if(StringUtils.isNotEmpty(title)){
                autoWriteConfig.setTitle(title);
            }
            autoWriteConfig.setOutputStream(outputStream);
            AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(autoWriteConfig);
            autoExcelWriter.write(headers,data);
            autoExcelWriter.close();
        } catch (Exception e) {
            throw new AxolotlWriteException("写入时发生异常："+e.getMessage());
        }
    }


    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,title,outputStream,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,title,outputStream,styleRender,sheetName,fontName,false,autoCatchColumnLength,false);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName){
        autoWriteToExcel(headers,data,title,outputStream,styleRender,sheetName,fontName,false,false,false);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String fontName){
        autoWriteToExcel(headers,data,title,outputStream,styleRender,null,fontName,false,false,false);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,title,outputStream,styleRender,null,null,false,autoCatchColumnLength,false);
    }
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender){
        autoWriteToExcel(headers,data,title,outputStream,styleRender,null,null,false,false,false);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber);
    }
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,title,outputStream,null,sheetName,fontName,false,autoCatchColumnLength,false);
    }
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName){
        autoWriteToExcel(headers,data,title,outputStream,null,sheetName,fontName,false,false,false);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName){
        autoWriteToExcel(headers,data,title,outputStream,null,null,fontName,false,false,false);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,title,outputStream,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,title,outputStream,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,title,outputStream,null,null,null,false,autoCatchColumnLength,false);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream){
        autoWriteToExcel(headers,data,title,outputStream,null,null,null,false,false,false);
    }

    public static void autoWriteToExcel(List<Header> headers, List<?> data, OutputStream outputStream){
        autoWriteToExcel(headers,data,null,outputStream,null,null,null,false,false,false);
    }

    public static void autoWriteToExcel(List<?> data, OutputStream outputStream){
        autoWriteToExcel(null,data,null,outputStream,null,null,null,false,false,false);
    }


    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoWriteConfig 自动写入配置
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, AutoWriteConfig autoWriteConfig){
        try {
            AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(autoWriteConfig);
            autoExcelWriter.write(headers,data);
            autoExcelWriter.close();
        } catch (Exception e) {
            throw new AxolotlWriteException("写入时发生异常："+e.getMessage());
        }
    }


    public static void autoWriteToExcelMultiSheet(OutputStream outputStream, SheetDataPackage... sheetDataPackage){
        try {
            if(sheetDataPackage == null || sheetDataPackage.length == 0){
                throw new AxolotlWriteException("写入配置不能为空");
            }
            AutoWriteConfig config = new AutoWriteConfig();
            config.setOutputStream(outputStream);
            AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(config);
            for (SheetDataPackage sheet : sheetDataPackage) {
                AutoWriteConfig autoWriteConfig = sheet.getAutoWriteConfig();
                List<Header> headers = sheet.getHeaders();
                List<?> data = sheet.getData();
                if(autoWriteConfig == null){
                    throw new AxolotlWriteException("写入配置不能为空");
                }
                autoExcelWriter.switchSheet(autoWriteConfig.getSheetIndex());
                copyAutoWriteConfig(autoWriteConfig,config);
                autoExcelWriter.write(headers,data);
            }
            autoExcelWriter.close();
        } catch (Exception e) {
            throw new AxolotlWriteException("写入时发生异常："+e.getMessage());
        }
    }

    private static void copyAutoWriteConfig(AutoWriteConfig source,AutoWriteConfig target){
        source.setOutputStream(target.getOutputStream());
        for (Integer sheetIndex : source.getCalculateColumnIndexes().keySet()) {
            if(target.getCalculateColumnIndexes().containsKey(sheetIndex)){
                Set<Integer> old = target.getCalculateColumnIndexes().get(sheetIndex);
                source.getCalculateColumnIndexes().get(sheetIndex).addAll(old);
            }
        }
        for (Integer sheetIndex : target.getCalculateColumnIndexes().keySet()) {
            if(!source.getCalculateColumnIndexes().containsKey(sheetIndex)){
                Set<Integer> old = target.getCalculateColumnIndexes().get(sheetIndex);
                source.getCalculateColumnIndexes().put(sheetIndex,old);
            }
        }
        source.getSpecialRowHeightMapping().putAll(target.getSpecialRowHeightMapping());
        try {
            BeanUtils.copyProperties(target,source);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    /**
     * 构建写入配置
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoCatchColumnLength, boolean autoInsertSerialNumber){
        AutoWriteConfig autoWriteConfig = new AutoWriteConfig();
        autoWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,autoInsertSerialNumber);
        autoWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING,autoInsertTotalInEnding);
        autoWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH,autoCatchColumnLength);
        if(styleRender != null){
            autoWriteConfig.setStyleRender(styleRender);
        }
        if(StringUtils.isNotEmpty(sheetName)){
            autoWriteConfig.setSheetName(sheetName);
        }
        if(StringUtils.isNotEmpty(fontName)){
            autoWriteConfig.setFontName(fontName);
        }
        if(StringUtils.isNotEmpty(title)){
            autoWriteConfig.setTitle(title);
        }
        autoWriteConfig.setSheetIndex(sheetIndex);
        SheetDataPackage sheetDataPackage = new SheetDataPackage();
        sheetDataPackage.setAutoWriteConfig(autoWriteConfig);
        sheetDataPackage.setHeaders(headers != null ? new ArrayList<>(headers) : null);
        sheetDataPackage.setData(data != null ? new ArrayList<>(data) : null);
        return sheetDataPackage;
    }


    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,styleRender,sheetName,fontName,false,autoCatchColumnLength,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,styleRender,sheetName,fontName,false,false,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String fontName){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,fontName,false,false,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoCatchColumnLength){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,null,false,autoCatchColumnLength,false);
    }
    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,null,false,false,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber);
    }
    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoCatchColumnLength){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,null,sheetName,fontName,false,autoCatchColumnLength,false);
    }
    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,null,sheetName,fontName,false,false,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,null,null,fontName,false,false,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoCatchColumnLength){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,null,null,null,false,autoCatchColumnLength,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title){
        return buildWriteSheetInfo(sheetIndex,headers,data,title,null,null,null,false,false,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data){
        return buildWriteSheetInfo(sheetIndex,headers,data,null,null,null,null,false,false,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<?> data){
        return buildWriteSheetInfo(sheetIndex,null,data,null,null,null,null,false,false,false);
    }

    public static SheetDataPackage buildWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, AutoWriteConfig autoWriteConfig){
        AutoWriteConfig config = new AutoWriteConfig();
        try {
            BeanUtils.copyProperties(config,autoWriteConfig);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        config.setSheetIndex(sheetIndex);
        SheetDataPackage sheetDataPackage = new SheetDataPackage();
        sheetDataPackage.setAutoWriteConfig(config);
        sheetDataPackage.setHeaders(headers != null ? new ArrayList<>(headers) : null);
        sheetDataPackage.setData(data != null ? new ArrayList<>(data) : null);
        return sheetDataPackage;
    }

}

package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.excel.writer.support.AutoSheetDataPackage;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlTemplateExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.TemplateWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.TemplateSheetDataPackage;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.formula.functions.T;

import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * 快拆工具类
 * @author 张智凯
 * @version 1.0.15
 */
public class AxolotlFaster {

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
        TemplateWriteConfig templateWriteConfig = new TemplateWriteConfig();
        templateWriteConfig.setOutputStream(outputStream);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW,templateShiftWriteRow);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL,templateNullValueWithTemplateFill);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL,templateNonTemplateCellFill);
        AxolotlTemplateExcelWriter templateExcelWriter = Axolotls.getTemplateExcelWriter(template, templateWriteConfig);
        AxolotlWriteResult result = templateExcelWriter.write(fixMapping, datas);
        templateExcelWriter.close();
        return result;
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
       return writeToTemplate(template,outputStream,fixMapping,datas,true,templateShiftWriteRow,templateNonTemplateCellFill);
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill){
        return writeToTemplate(template,outputStream,fixMapping,datas,templateNullValueWithTemplateFill,true,true);
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas){
        return writeToTemplate(template,outputStream,fixMapping,datas,true,true,true);
    }


    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
        return writeToTemplate(template,outputStream,null,datas,true,templateShiftWriteRow,templateNonTemplateCellFill);
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas, boolean templateNullValueWithTemplateFill){
        return writeToTemplate(template,outputStream,null,datas,templateNullValueWithTemplateFill,true,true);
    }

    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas){
        return writeToTemplate(template,outputStream,null,datas,true,true,true);
    }

    /**
     * 模板写入
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
     * 一次写入多个sheet
     * @param template 模板
     * @param outputStream 输出流
     * @param sheetDataPackage 单个sheet配置
     */
    public static void templateWriteToExcelMultiSheet(File template, OutputStream outputStream, TemplateSheetDataPackage... sheetDataPackage){
        try {
            if(sheetDataPackage == null || sheetDataPackage.length == 0){
                throw new AxolotlWriteException("写入配置不能为空");
            }
            TemplateWriteConfig config = new TemplateWriteConfig();
            config.setOutputStream(outputStream);
            AxolotlTemplateExcelWriter templateExcelWriter = Axolotls.getTemplateExcelWriter(template, config);
            for (TemplateSheetDataPackage sheet : sheetDataPackage) {
                TemplateWriteConfig templateWriteConfig = sheet.getTemplateWriteConfig();
                Map<String, ?> fixMapping = sheet.getFixMapping();
                List<?> datas = sheet.getDatas();
                if(templateWriteConfig == null){
                    throw new AxolotlWriteException("写入配置不能为空");
                }
                templateExcelWriter.switchSheet(templateWriteConfig.getSheetIndex());
                copyTemplateWriteConfig(templateWriteConfig,config);
                templateExcelWriter.write(fixMapping,datas);
            }
            templateExcelWriter.close();
        } catch (Exception e) {
            throw new AxolotlWriteException("写入时发生异常："+e.getMessage());
        }
    }

    private static void copyTemplateWriteConfig(TemplateWriteConfig source,TemplateWriteConfig target){
        source.setOutputStream(target.getOutputStream());
        source.getDictionaryMapping().putAll(target.getDictionaryMapping());
        try {
            BeanUtils.copyProperties(target,source);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 构建模板写入配置
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
        TemplateWriteConfig templateWriteConfig = new TemplateWriteConfig();
        templateWriteConfig.setSheetIndex(sheetIndex);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW,templateShiftWriteRow);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL,templateNullValueWithTemplateFill);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL,templateNonTemplateCellFill);
        TemplateSheetDataPackage sheetInfo = new TemplateSheetDataPackage();
        sheetInfo.setTemplateWriteConfig(templateWriteConfig);
        sheetInfo.setFixMapping(fixMapping != null ? new HashMap<>(fixMapping) : null);
        sheetInfo.setDatas(datas != null ? new ArrayList<>(datas) : null);
        return sheetInfo;
    }

    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,true,templateShiftWriteRow,templateNonTemplateCellFill);
    }

    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,templateNullValueWithTemplateFill,true,true);
    }

    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,true,true,true);
    }


    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
        return buildTemplateWriteSheetInfo(sheetIndex,null,datas,true,templateShiftWriteRow,templateNonTemplateCellFill);
    }

    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, List<?> datas, boolean templateNullValueWithTemplateFill){
        return buildTemplateWriteSheetInfo(sheetIndex,null,datas,templateNullValueWithTemplateFill,true,true);
    }

    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, List<?> datas){
        return buildTemplateWriteSheetInfo(sheetIndex,null,datas,true,true,true);
    }


    /**
     * 构建模板写入配置
     * @param sheetIndex 表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateWriteConfig 模板写入配置
     * @return
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, TemplateWriteConfig templateWriteConfig){
        TemplateWriteConfig config = new TemplateWriteConfig();
        try {
            BeanUtils.copyProperties(config,templateWriteConfig);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        config.setSheetIndex(sheetIndex);
        TemplateSheetDataPackage sheetInfo = new TemplateSheetDataPackage();
        sheetInfo.setTemplateWriteConfig(config);
        sheetInfo.setFixMapping(fixMapping != null ? new HashMap<>(fixMapping) : null);
        sheetInfo.setDatas(datas != null ? new ArrayList<>(datas) : null);
        return sheetInfo;
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


    /**
     * 自动写入多个sheet
     * @param outputStream 输出流
     * @param sheetDataPackage 单个sheet配置
     */
    public static void autoWriteToExcelMultiSheet(OutputStream outputStream, AutoSheetDataPackage... sheetDataPackage){
        try {
            if(sheetDataPackage == null || sheetDataPackage.length == 0){
                throw new AxolotlWriteException("写入配置不能为空");
            }
            AutoWriteConfig config = new AutoWriteConfig();
            config.setOutputStream(outputStream);
            AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(config);
            for (AutoSheetDataPackage sheet : sheetDataPackage) {
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
        source.getDictionaryMapping().putAll(target.getDictionaryMapping());
        try {
            BeanUtils.copyProperties(target,source);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    /**
     * 构建自动写入配置
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
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoCatchColumnLength, boolean autoInsertSerialNumber){
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
        AutoSheetDataPackage sheetDataPackage = new AutoSheetDataPackage();
        sheetDataPackage.setAutoWriteConfig(autoWriteConfig);
        sheetDataPackage.setHeaders(headers != null ? new ArrayList<>(headers) : null);
        sheetDataPackage.setData(data != null ? new ArrayList<>(data) : null);
        return sheetDataPackage;
    }


    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,styleRender,sheetName,fontName,false,autoCatchColumnLength,false);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,styleRender,sheetName,fontName,false,false,false);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String fontName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,fontName,false,false,false);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,null,false,autoCatchColumnLength,false);
    }
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,styleRender,null,null,false,false,false);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber);
    }
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,null,sheetName,fontName,false,autoCatchColumnLength,false);
    }
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,null,sheetName,fontName,false,false,false);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,null,null,fontName,false,false,false);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,null,null,null,false,autoCatchColumnLength,false);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,title,null,null,null,false,false,false);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,null,null,null,false,false,false);
    }

    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<?> data){
        return buildAutoWriteSheetInfo(sheetIndex,null,data,null,null,null,null,false,false,false);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex sheet索引
     * @param headers 表头
     * @param data 数据
     * @param autoWriteConfig 写入配置
     * @return
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, AutoWriteConfig autoWriteConfig){
        AutoWriteConfig config = new AutoWriteConfig();
        try {
            BeanUtils.copyProperties(config,autoWriteConfig);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        config.setSheetIndex(sheetIndex);
        AutoSheetDataPackage sheetDataPackage = new AutoSheetDataPackage();
        sheetDataPackage.setAutoWriteConfig(config);
        sheetDataPackage.setHeaders(headers != null ? new ArrayList<>(headers) : null);
        sheetDataPackage.setData(data != null ? new ArrayList<>(data) : null);
        return sheetDataPackage;
    }

}

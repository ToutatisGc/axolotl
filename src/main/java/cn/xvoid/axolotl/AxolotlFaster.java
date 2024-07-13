package cn.xvoid.axolotl;

import cn.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.xvoid.axolotl.excel.reader.AxolotlStreamExcelReader;
import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.xvoid.axolotl.excel.reader.support.stream.AxolotlExcelStream;
import cn.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import cn.xvoid.axolotl.excel.writer.AxolotlTemplateExcelWriter;
import cn.xvoid.axolotl.excel.writer.TemplateWriteConfig;
import cn.xvoid.axolotl.excel.writer.components.widgets.Header;
import cn.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.xvoid.axolotl.excel.writer.support.AutoSheetDataPackage;
import cn.xvoid.axolotl.excel.writer.support.TemplateSheetDataPackage;
import cn.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
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
     * @param dict 字典值
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill, Map<String,Object> dict){
        TemplateWriteConfig templateWriteConfig = new TemplateWriteConfig();
        templateWriteConfig.setOutputStream(outputStream);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW,templateShiftWriteRow);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL,templateNullValueWithTemplateFill);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL,templateNonTemplateCellFill);
        if(dict != null){
            for (String fieldName : dict.keySet()) {
                if(fieldName != null){
                    Object dictParam = dict.get(fieldName);
                    if(dictParam != null){
                        if(dictParam instanceof List<?>){
                            templateWriteConfig.setDict(templateWriteConfig.getSheetIndex(),fieldName, new ArrayList<>((List<?>) dictParam));
                        }else if(dictParam instanceof Map<?,?>){
                            templateWriteConfig.setDict(templateWriteConfig.getSheetIndex(),fieldName, new HashMap<>((Map<String, String>) dictParam));
                        }
                    }
                }
            }
        }
        AxolotlTemplateExcelWriter templateExcelWriter = Axolotls.getTemplateExcelWriter(template, templateWriteConfig);
        AxolotlWriteResult result = templateExcelWriter.write(fixMapping, datas);
        templateExcelWriter.close();
        return result;
    }

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
        return writeToTemplate(template,outputStream,fixMapping,datas,templateNullValueWithTemplateFill,templateShiftWriteRow,templateNonTemplateCellFill,null);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
       return writeToTemplate(template,outputStream,fixMapping,datas,true,templateShiftWriteRow,templateNonTemplateCellFill,null);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill){
        return writeToTemplate(template,outputStream,fixMapping,datas,templateNullValueWithTemplateFill,true,true,null);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas){
        return writeToTemplate(template,outputStream,fixMapping,datas,true,true,true,null);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param datas 列表数据
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
        return writeToTemplate(template,outputStream,null,datas,true,templateShiftWriteRow,templateNonTemplateCellFill,null);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas, boolean templateNullValueWithTemplateFill){
        return writeToTemplate(template,outputStream,null,datas,templateNullValueWithTemplateFill,true,true,null);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param datas 列表数据
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas){
        return writeToTemplate(template,outputStream,null,datas,true,true,true,null);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @param dict 字典值
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill, Map<String,Object> dict){
        return writeToTemplate(template,outputStream,fixMapping,datas,true,templateShiftWriteRow,templateNonTemplateCellFill,dict);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @param dict 字典值
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill, Map<String,Object> dict){
        return writeToTemplate(template,outputStream,fixMapping,datas,templateNullValueWithTemplateFill,true,true,dict);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param dict 字典值
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, Map<String,?> fixMapping, List<?> datas, Map<String,Object> dict){
        return writeToTemplate(template,outputStream,fixMapping,datas,true,true,true,dict);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param datas 列表数据
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @param dict 字典值
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill, Map<String,Object> dict){
        return writeToTemplate(template,outputStream,null,datas,true,templateShiftWriteRow,templateNonTemplateCellFill,dict);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @param dict 字典值
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas, boolean templateNullValueWithTemplateFill, Map<String,Object> dict){
        return writeToTemplate(template,outputStream,null,datas,templateNullValueWithTemplateFill,true,true,dict);
    }

    /**
     * 模板写入
     * @param template 模板文件
     * @param outputStream 输出流
     * @param datas 列表数据
     * @param dict 字典值
     * @return 写入结果
     */
    public static AxolotlWriteResult writeToTemplate(File template, OutputStream outputStream, List<?> datas, Map<String,Object> dict){
        return writeToTemplate(template,outputStream,null,datas,true,true,true,dict);
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

    /**
     * 配置信息载入
     * @param source 需要载入的配置信息
     * @param target 载入的目标配置类
     */
    private static void copyTemplateWriteConfig(TemplateWriteConfig source,TemplateWriteConfig target){
        source.setOutputStream(target.getOutputStream());
        //合并字典映射信息
        target.getDictionaryMapping().putAll(source.getDictionaryMapping());
        source.setDictionaryMapping(target.getDictionaryMapping());
        try {
            BeanUtils.copyProperties(target,source);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @param dict 字典值列表(属性名,字典值参数)
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill, Map<String,Object> dict){
        TemplateWriteConfig templateWriteConfig = new TemplateWriteConfig();
        templateWriteConfig.setSheetIndex(sheetIndex);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW,templateShiftWriteRow);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL,templateNullValueWithTemplateFill);
        templateWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL,templateNonTemplateCellFill);
        if(dict != null){
            for (String fieldName : dict.keySet()) {
                if(fieldName != null){
                    Object dictParam = dict.get(fieldName);
                    if(dictParam != null){
                        if(dictParam instanceof List<?>){
                            templateWriteConfig.setDict(sheetIndex,fieldName, new ArrayList<>((List<?>) dictParam));
                        }else if(dictParam instanceof Map<?,?>){
                            templateWriteConfig.setDict(sheetIndex,fieldName, new HashMap<>((Map<String, String>) dictParam));
                        }
                    }
                }
            }
        }
        TemplateSheetDataPackage sheetInfo = new TemplateSheetDataPackage();
        sheetInfo.setTemplateWriteConfig(templateWriteConfig);
        sheetInfo.setFixMapping(fixMapping != null ? new HashMap<>(fixMapping) : null);
        sheetInfo.setDatas(datas != null ? new ArrayList<>(datas) : null);
        return sheetInfo;
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,templateNullValueWithTemplateFill,templateShiftWriteRow,templateNonTemplateCellFill,null);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,true,templateShiftWriteRow,templateNonTemplateCellFill,null);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,templateNullValueWithTemplateFill,true,true,null);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,true,true,true,null);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param datas 列表数据
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill){
        return buildTemplateWriteSheetInfo(sheetIndex,null,datas,true,templateShiftWriteRow,templateNonTemplateCellFill,null);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, List<?> datas, boolean templateNullValueWithTemplateFill){
        return buildTemplateWriteSheetInfo(sheetIndex,null,datas,templateNullValueWithTemplateFill,true,true,null);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param datas 列表数据
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, List<?> datas){
        return buildTemplateWriteSheetInfo(sheetIndex,null,datas,true,true,true,null);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @param dict 字典值列表(属性名,字典值参数)
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill, Map<String,Object> dict){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,true,templateShiftWriteRow,templateNonTemplateCellFill,dict);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @param dict 字典值列表(属性名,字典值参数)
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, boolean templateNullValueWithTemplateFill, Map<String,Object> dict){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,templateNullValueWithTemplateFill,true,true,dict);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param dict 字典值列表(属性名,字典值参数)
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, Map<String,?> fixMapping, List<?> datas, Map<String,Object> dict){
        return buildTemplateWriteSheetInfo(sheetIndex,fixMapping,datas,true,true,true,dict);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param datas 列表数据
     * @param templateShiftWriteRow  特性 ExcelWritePolicy.TEMPLATE_SHIFT_WRITE_ROW  处理#{}占位符时，新增行是否创建，反之则直接使用现有行
     * @param templateNonTemplateCellFill  特性  ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL  处理#{}占位符时，与占位符同行的其他单元格数据是否也渲染到新增行上
     * @param dict 字典值列表(属性名,字典值参数)
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, List<?> datas, boolean templateShiftWriteRow, boolean templateNonTemplateCellFill, Map<String,Object> dict){
        return buildTemplateWriteSheetInfo(sheetIndex,null,datas,true,templateShiftWriteRow,templateNonTemplateCellFill,dict);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param datas 列表数据
     * @param templateNullValueWithTemplateFill  特性 ExcelWritePolicy.TEMPLATE_NULL_VALUE_WITH_TEMPLATE_FILL  当占位符值为空时是否将整个单元格赋值为空
     * @param dict 字典值列表(属性名,字典值参数)
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, List<?> datas, boolean templateNullValueWithTemplateFill, Map<String,Object> dict){
        return buildTemplateWriteSheetInfo(sheetIndex,null,datas,templateNullValueWithTemplateFill,true,true,dict);
    }

    /**
     * 构建模板写入配置
     * @param sheetIndex 工作表索引
     * @param datas 列表数据
     * @param dict 字典值列表(属性名,字典值参数)
     * @return 写入结果
     */
    public static TemplateSheetDataPackage buildTemplateWriteSheetInfo(int sheetIndex, List<?> datas, Map<String,Object> dict){
        return buildTemplateWriteSheetInfo(sheetIndex,null,datas,true,true,true,dict);
    }


    /**
     * 构建模板写入配置
     * @param sheetIndex 表索引
     * @param fixMapping 引用字段
     * @param datas 列表数据
     * @param templateWriteConfig 模板写入配置
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
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoCatchColumnLength, boolean autoInsertSerialNumber, Map<String,Object> dict){
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
            if(dict != null){
                for (String fieldName : dict.keySet()) {
                    if(fieldName != null){
                        Object dictParam = dict.get(fieldName);
                        if(dictParam != null){
                            if(dictParam instanceof List<?>){
                                autoWriteConfig.setDict(autoWriteConfig.getSheetIndex(),fieldName, new ArrayList<>((List<?>) dictParam));
                            }else if(dictParam instanceof Map<?,?>){
                                autoWriteConfig.setDict(autoWriteConfig.getSheetIndex(),fieldName, new HashMap<>((Map<String, String>) dictParam));
                            }
                        }
                    }
                }
            }
            if(autoColumnWidthRatio != null){
                autoWriteConfig.setAutoColumnWidthRatio(autoColumnWidthRatio);
            }
            autoWriteConfig.setOutputStream(outputStream);
            AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(autoWriteConfig);
            autoExcelWriter.write(headers,data);
            autoExcelWriter.close();
        } catch (Exception e) {
            throw new AxolotlWriteException("写入时发生异常："+e.getMessage());
        }
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
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoCatchColumnLength, boolean autoInsertSerialNumber, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
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
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
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
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
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
     * @param autoCatchColumnLength  自动列宽
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,sheetName,fontName,false,autoCatchColumnLength,false,null);
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
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,sheetName,fontName,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param fontName 字体名称
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String fontName){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,fontName,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoCatchColumnLength  自动列宽
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,null,false,autoCatchColumnLength,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,null,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,null,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,null,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,null,title,outputStream,null,sheetName,fontName,false,autoCatchColumnLength,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName){
        autoWriteToExcel(headers,data,null,title,outputStream,null,sheetName,fontName,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param fontName 字体名称
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,fontName,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param autoCatchColumnLength  自动列宽
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,null,false,autoCatchColumnLength,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,null,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param outputStream 输出流
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, OutputStream outputStream){
        autoWriteToExcel(headers,data,null,null,outputStream,null,null,null,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param data 数据
     * @param outputStream 输出流
     */
    public static void autoWriteToExcel(List<?> data, OutputStream outputStream){
        autoWriteToExcel(null,data,null,null,outputStream,null,null,null,false,false,false,null);
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
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
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
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,sheetName,fontName,false,autoCatchColumnLength,false,dict);
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
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,sheetName,fontName,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param fontName 字体名称
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, String fontName, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,fontName,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,null,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, ExcelStyleRender styleRender, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,styleRender,null,null,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,null,sheetName,fontName,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, String sheetName, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,null,sheetName,fontName,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param fontName 字体名称
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, String fontName, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,fontName,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,null,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param outputStream 输出流
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, String title, OutputStream outputStream, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,title,outputStream,null,null,null,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param outputStream 输出流
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, OutputStream outputStream, Map<String,Object> dict){
        autoWriteToExcel(headers,data,null,null,outputStream,null,null,null,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param data 数据
     * @param outputStream 输出流
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<?> data, OutputStream outputStream, Map<String,Object> dict){
        autoWriteToExcel(null,data,null,null,outputStream,null,null,null,false,false,false,dict);
    }


    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoCatchColumnLength, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,sheetName,fontName,false,autoCatchColumnLength,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,sheetName,fontName,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param fontName 字体名称
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String fontName){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,fontName,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoCatchColumnLength  自动列宽
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,null,false,autoCatchColumnLength,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,null,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,sheetName,fontName,false,autoCatchColumnLength,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName, String sheetName){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,sheetName,fontName,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param fontName 字体名称
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,fontName,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param autoCatchColumnLength  自动列宽
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, boolean autoCatchColumnLength){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,null,false,autoCatchColumnLength,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,null,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param outputStream 输出流
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, OutputStream outputStream){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,null,outputStream,null,null,null,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param outputStream 输出流
     */
    public static void autoWriteToExcel(List<?> data, Double autoColumnWidthRatio, OutputStream outputStream){
        autoWriteToExcel(null,data,autoColumnWidthRatio,null,outputStream,null,null,null,false,false,false,null);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,sheetName,fontName,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String sheetName, String fontName, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,sheetName,fontName,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param fontName 字体名称
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, String fontName, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,fontName,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,null,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param styleRender 渲染器
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, ExcelStyleRender styleRender, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,styleRender,null,null,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName, String sheetName, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,sheetName,fontName,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName, String sheetName, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,sheetName,fontName,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param fontName 字体名称
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, String fontName, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,fontName,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, boolean autoCatchColumnLength, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,null,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param outputStream 输出流
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, OutputStream outputStream, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,title,outputStream,null,null,null,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param outputStream 输出流
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<Header> headers, List<?> data, Double autoColumnWidthRatio, OutputStream outputStream, Map<String,Object> dict){
        autoWriteToExcel(headers,data,autoColumnWidthRatio,null,outputStream,null,null,null,false,false,false,dict);
    }

    /**
     * 自动写入 Excel
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param outputStream 输出流
     * @param dict 字典值
     */
    public static void autoWriteToExcel(List<?> data, Double autoColumnWidthRatio, OutputStream outputStream, Map<String,Object> dict){
        autoWriteToExcel(null,data,autoColumnWidthRatio,null,outputStream,null,null,null,false,false,false,dict);
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

    /**
     * 配置信息载入
     * @param source 需要载入的配置信息
     * @param target 载入的目标配置类
     */
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
        target.getSpecialRowHeightMapping().putAll(source.getSpecialRowHeightMapping());
        target.getDictionaryMapping().putAll(source.getDictionaryMapping());
        source.setSpecialRowHeightMapping(target.getSpecialRowHeightMapping());
        source.setDictionaryMapping(target.getDictionaryMapping());
        try {
            BeanUtils.copyProperties(target,source);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoCatchColumnLength, boolean autoInsertSerialNumber, Map<String,Object> dict){
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
        if(dict != null){
            for (String fieldName : dict.keySet()) {
                if(fieldName != null){
                    Object dictParam = dict.get(fieldName);
                    if(dictParam != null){
                        if(dictParam instanceof List<?>){
                            autoWriteConfig.setDict(sheetIndex,fieldName, new ArrayList<>((List<?>) dictParam));
                        }else if(dictParam instanceof Map<?,?>){
                            autoWriteConfig.setDict(sheetIndex,fieldName, new HashMap<>((Map<String, String>) dictParam));
                        }
                    }
                }
            }
        }
        if(autoColumnWidthRatio != null){
            autoWriteConfig.setAutoColumnWidthRatio(autoColumnWidthRatio);
        }
        autoWriteConfig.setSheetIndex(sheetIndex);
        AutoSheetDataPackage sheetDataPackage = new AutoSheetDataPackage();
        sheetDataPackage.setAutoWriteConfig(autoWriteConfig);
        sheetDataPackage.setHeaders(headers != null ? new ArrayList<>(headers) : null);
        sheetDataPackage.setData(data != null ? new ArrayList<>(data) : null);
        return sheetDataPackage;
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoCatchColumnLength, boolean autoInsertSerialNumber, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
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
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,sheetName,fontName,false,autoCatchColumnLength,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,sheetName,fontName,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param fontName 字体名称
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String fontName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,fontName,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoCatchColumnLength  自动列宽
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,null,false,autoCatchColumnLength,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,null,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,sheetName,fontName,false,autoCatchColumnLength,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,sheetName,fontName,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param fontName 字体名称
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,fontName,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param autoCatchColumnLength  自动列宽
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,null,false,autoCatchColumnLength,false,null);
    }

    /**
     * 构建自动写入配置
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,null,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,null,null,null,null,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param data 数据
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<?> data){
        return buildAutoWriteSheetInfo(sheetIndex,null,data,null,null,null,null,null,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,sheetName,fontName,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String sheetName, String fontName, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,sheetName,fontName,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param fontName 字体名称
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, String fontName, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,fontName,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据

     * @param title 标题
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,null,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param styleRender 渲染器
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, ExcelStyleRender styleRender, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,styleRender,null,null,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,sheetName,fontName,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, String sheetName, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,sheetName,fontName,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param fontName 字体名称
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, String fontName, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,fontName,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,null,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param title 标题
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, String title, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,title,null,null,null,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,null,null,null,null,null,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param data 数据
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<?> data, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,null,data,null,null,null,null,null,false,false,false,dict);
    }


    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoCatchColumnLength, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,sheetName,fontName,false,autoCatchColumnLength,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String sheetName, String fontName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,sheetName,fontName,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param fontName 字体名称
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String fontName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,fontName,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoCatchColumnLength  自动列宽
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,null,false,autoCatchColumnLength,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,null,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName, String sheetName, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,sheetName,fontName,false,autoCatchColumnLength,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName, String sheetName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,sheetName,fontName,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param fontName 字体名称
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,fontName,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param autoCatchColumnLength  自动列宽
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, boolean autoCatchColumnLength){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,null,false,autoCatchColumnLength,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,null,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,null,null,null,null,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<?> data, Double autoColumnWidthRatio){
        return buildAutoWriteSheetInfo(sheetIndex,null,data,autoColumnWidthRatio,null,null,null,null,false,false,false,null);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String sheetName, String fontName, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,sheetName,fontName,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String sheetName, String fontName, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,sheetName,fontName,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param fontName 字体名称
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, String fontName, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,fontName,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,null,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param styleRender 渲染器
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, ExcelStyleRender styleRender, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,styleRender,null,null,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,sheetName,fontName,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName, String sheetName, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,sheetName,fontName,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName, String sheetName, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,sheetName,fontName,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param sheetName sheet名称
     * @param fontName 字体名称
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName, String sheetName, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,sheetName,fontName,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param fontName 字体名称
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, String fontName, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,fontName,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoCatchColumnLength  自动列宽
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,null,autoInsertTotalInEnding,autoCatchColumnLength,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param autoInsertTotalInEnding 结尾添加合计
     * @param autoInsertSerialNumber 第一行添加序号
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, boolean autoInsertTotalInEnding, boolean autoInsertSerialNumber, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,null,autoInsertTotalInEnding,false,autoInsertSerialNumber,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param autoCatchColumnLength  自动列宽
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, boolean autoCatchColumnLength, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,null,false,autoCatchColumnLength,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param title 标题
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, String title, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,title,null,null,null,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param headers 表头
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<Header> headers, List<?> data, Double autoColumnWidthRatio, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,headers,data,autoColumnWidthRatio,null,null,null,null,false,false,false,dict);
    }

    /**
     * 构建自动写入配置
     * @param sheetIndex 工作表索引
     * @param data 数据
     * @param autoColumnWidthRatio 自动列宽比例
     * @param dict 字典值列表(属性名,字典值参数)
     */
    public static AutoSheetDataPackage buildAutoWriteSheetInfo(int sheetIndex, List<?> data, Double autoColumnWidthRatio, Map<String,Object> dict){
        return buildAutoWriteSheetInfo(sheetIndex,null,data,autoColumnWidthRatio,null,null,null,null,false,false,false,dict);
    }


    /**
     * 构建自动写入配置
     * @param sheetIndex sheet索引
     * @param headers 表头
     * @param data 数据
     * @param autoWriteConfig 写入配置
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

    /**
     * 获取无泛型Excel读取器
     * @param excelFile excel文件
     */
    public static AxolotlExcelReader<?> getExcelReader(File excelFile){
        return Axolotls.getExcelReader(excelFile);
    }

    /**
     * 获取无泛型Excel读取器
     * @param inputStream excel文件流
     */
    public static AxolotlExcelReader<?> getExcelReader(InputStream inputStream){
        return Axolotls.getExcelReader(inputStream);
    }

    /**
     * 读取sheet为一个List
     * @param reader 读取器
     * @param config 读取配置
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, ReaderConfig<T> config){
        if(reader == null){
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR, "读取器不能为空");
        }
        return reader.readSheetData(config);
    }


    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        if(reader == null){
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR, "读取器不能为空");
        }
        ReaderConfig<T> config = new ReaderConfig<>();
        if(sheetIndex != null){
            config.setSheetIndex(sheetIndex);
        }
        if(initialRowPositionOffset != null){
            config.setInitialRowPositionOffset(initialRowPositionOffset);
        }
        if(startRowIndex != null){
            config.setStartIndex(startRowIndex);
        }
        if(endRowIndex != null){
            config.setEndIndex(endRowIndex);
        }
        if(startColumnIndex != null){
            config.setSheetColumnEffectiveRangeStart(startColumnIndex);
        }
        if(endColumnIndex != null){
            config.setSheetColumnEffectiveRangeEnd(endColumnIndex);
        }
        if(searchHeaderMaxRows != null){
            config.setSearchHeaderMaxRows(searchHeaderMaxRows);
        }
        if(dict != null){
            for (String fieldName : dict.keySet()) {
                if(fieldName != null){
                    Object dictParam = dict.get(fieldName);
                    if(dictParam != null){
                        if(dictParam instanceof List<?>){
                            config.setDict(config.getSheetIndex(),fieldName, new ArrayList<>((List<?>) dictParam));
                        }else if(dictParam instanceof Map<?,?>){
                            config.setDict(config.getSheetIndex(),fieldName, new HashMap<>((Map<String, String>) dictParam));
                        }
                    }
                }
            }
        }
        if(validateReadRowData != null){
            config.setBooleanReadPolicy(ExcelReadPolicy.VALIDATE_READ_ROW_DATA,validateReadRowData);
        }
        if(trimCellValue != null){
            config.setBooleanReadPolicy(ExcelReadPolicy.TRIM_CELL_VALUE,trimCellValue);
        }
        config.setCastClass(clazz);
        return reader.readSheetData(config);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Map<String,Object> dict){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,dict,null,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,null,null,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,null,null,null,null,null,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,null,null,null,null,null,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Map<String,Object> dict){
        return readSheetAsList(reader,clazz,sheetIndex,null,null,null,null,null,null,dict,null,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,null,null,null,null,null,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,null,null,null,null,null,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex){
        return readSheetAsList(reader,clazz,sheetIndex,null,null,null,null,null,null,null,null,null);
    }


    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Map<String,Object> dict){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,dict,null,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,null,null,null);
    }


    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,dict,null,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex){
        return readSheetAsList(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,null,null,null);
    }


    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Map<String,Object> dict){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,dict,null,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,null,null,null);
    }


    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,dict,null,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个List
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @return 读取结果  List集合
     */
    public static <T> List<T> readSheetAsList(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex){
        return readSheetAsList(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,null,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader 读取器
     * @param config 读取配置
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, ReaderConfig<T> config){
        if(reader == null){
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR, "读取器不能为空");
        }
        return reader.readSheetDataAsObject(config);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        if(reader == null){
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR, "读取器不能为空");
        }
        ReaderConfig<T> config = new ReaderConfig<>();
        if(sheetIndex != null){
            config.setSheetIndex(sheetIndex);
        }
        if(initialRowPositionOffset != null){
            config.setInitialRowPositionOffset(initialRowPositionOffset);
        }
        if(startRowIndex != null){
            config.setStartIndex(startRowIndex);
        }
        if(endRowIndex != null){
            config.setEndIndex(endRowIndex);
        }
        if(startColumnIndex != null){
            config.setSheetColumnEffectiveRangeStart(startColumnIndex);
        }
        if(endColumnIndex != null){
            config.setSheetColumnEffectiveRangeEnd(endColumnIndex);
        }
        if(searchHeaderMaxRows != null){
            config.setSearchHeaderMaxRows(searchHeaderMaxRows);
        }
        if(dict != null){
            for (String fieldName : dict.keySet()) {
                if(fieldName != null){
                    Object dictParam = dict.get(fieldName);
                    if(dictParam != null){
                        if(dictParam instanceof List<?>){
                            config.setDict(config.getSheetIndex(),fieldName, new ArrayList<>((List<?>) dictParam));
                        }else if(dictParam instanceof Map<?,?>){
                            config.setDict(config.getSheetIndex(),fieldName, new HashMap<>((Map<String, String>) dictParam));
                        }
                    }
                }
            }
        }
        if(validateReadRowData != null){
            config.setBooleanReadPolicy(ExcelReadPolicy.VALIDATE_READ_ROW_DATA,validateReadRowData);
        }
        if(trimCellValue != null){
            config.setBooleanReadPolicy(ExcelReadPolicy.TRIM_CELL_VALUE,trimCellValue);
        }
        config.setCastClass(clazz);
        return reader.readSheetDataAsObject(config);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Map<String,Object> dict){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,dict,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param searchHeaderMaxRows 按表头名绑定列时，读取表头最大行数 默认为10
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Integer searchHeaderMaxRows){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,searchHeaderMaxRows,null,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引

     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,null,null,null,null,null,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,null,null,null,null,null,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Map<String,Object> dict){
        return readSheetAsObject(reader,clazz,sheetIndex,null,null,null,null,null,null,dict,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,null,null,null,null,null,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,null,null,null,null,null,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex){
        return readSheetAsObject(reader,clazz,sheetIndex,null,null,null,null,null,null,null,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Map<String,Object> dict){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,dict,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,null,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,dict,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex){
        return readSheetAsObject(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,null,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Map<String,Object> dict){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,dict,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,null,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,dict,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,dict,null,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,null,validateReadRowData,null);
    }

    /**
     * 读取sheet为一个java对象
     * @param reader  读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @return 读取结果  java对象
     */
    public static <T> T readSheetAsObject(AxolotlExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex){
        return readSheetAsObject(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,null,null,null);
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
     * 使用流形式读取hseet
     * @param reader Excel流形式读取器
     * @param config 读取配置
     * @return 读取结果 流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, ReaderConfig<T> config){
        if(reader == null){
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR, "读取器不能为空");
        }
        return reader.dataIterator(config);
    }


    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        if(reader == null){
            throw new AxolotlWriteException("读取器不能为空");
        }
        ReaderConfig<T> config = new ReaderConfig<>();
        if(sheetIndex != null){
            config.setSheetIndex(sheetIndex);
        }
        if(initialRowPositionOffset != null){
            config.setInitialRowPositionOffset(initialRowPositionOffset);
        }
        if(startRowIndex != null){
            config.setStartIndex(startRowIndex);
        }
        if(endRowIndex != null){
            config.setEndIndex(endRowIndex);
        }
        if(startColumnIndex != null){
            config.setSheetColumnEffectiveRangeStart(startColumnIndex);
        }
        if(endColumnIndex != null){
            config.setSheetColumnEffectiveRangeEnd(endColumnIndex);
        }
        if(dict != null){
            for (String fieldName : dict.keySet()) {
                if(fieldName != null){
                    Object dictParam = dict.get(fieldName);
                    if(dictParam != null){
                        if(dictParam instanceof List<?>){
                            config.setDict(config.getSheetIndex(),fieldName, new ArrayList<>((List<?>) dictParam));
                        }else if(dictParam instanceof Map<?,?>){
                            config.setDict(config.getSheetIndex(),fieldName, new HashMap<>((Map<String, String>) dictParam));
                        }
                    }
                }
            }
        }
        if(validateReadRowData != null){
            config.setBooleanReadPolicy(ExcelReadPolicy.VALIDATE_READ_ROW_DATA,validateReadRowData);
        }
        if(trimCellValue != null){
            config.setBooleanReadPolicy(ExcelReadPolicy.TRIM_CELL_VALUE,trimCellValue);
        }
        config.setCastClass(clazz);
        return reader.dataIterator(config);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,dict,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,dict,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引

     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,null,null,null,null,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,null,null,null,null,null,dict,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Map<String,Object> dict){
        return readSheetUseStream(reader,clazz,sheetIndex,null,null,null,null,null,dict,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,null,null,null,null,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,null,null,null,null,null,null,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex){
        return readSheetUseStream(reader,clazz,sheetIndex,null,null,null,null,null,null,null,null);
    }


    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验

     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,dict,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Map<String,Object> dict){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,dict,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,null,null,null,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,dict,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,dict,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset  初始行偏移量
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer initialRowPositionOffset, Integer startColumnIndex, Integer endColumnIndex){
        return readSheetUseStream(reader,clazz,sheetIndex,initialRowPositionOffset,null,null,startColumnIndex,endColumnIndex,null,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,dict,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Map<String,Object> dict){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,dict,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,null,null,null,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,dict,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,dict,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param dict 字典值  Map(字段名,Map(字典名,字典值))  或   Map(字段名,List(Map或实体类{key:字典名,value:字典值}))
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Map<String,Object> dict){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,dict,null,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @param trimCellValue 是否开启单元格修整  开启后读取时将去掉单元格所有的空格和换行符
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData, Boolean trimCellValue){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,validateReadRowData,trimCellValue);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @param validateReadRowData 是否开启数据有效性校验
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex, Boolean validateReadRowData){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,validateReadRowData,null);
    }

    /**
     * 使用流形式读取sheet  适用于数据量较多的文件
     * @param reader  Excel流形式读取器
     * @param clazz  读取的Java类型
     * @param sheetIndex sheet索引
     * @param startRowIndex 读取范围：读取开始 行索引
     * @param endRowIndex 读取范围：读取结束 行索引
     * @param startColumnIndex 读取范围：读取开始 列索引
     * @param endColumnIndex 读取范围：读取开始 列索引
     * @return 读取结果  流形式的迭代器
     */
    public static <T> AxolotlExcelStream<T> readSheetUseStream(AxolotlStreamExcelReader<?> reader, Class<T> clazz, Integer sheetIndex, Integer startRowIndex, Integer endRowIndex, Integer startColumnIndex, Integer endColumnIndex){
        return readSheetUseStream(reader,clazz,sheetIndex,null,startRowIndex,endRowIndex,startColumnIndex,endColumnIndex,null,null,null);
    }




}

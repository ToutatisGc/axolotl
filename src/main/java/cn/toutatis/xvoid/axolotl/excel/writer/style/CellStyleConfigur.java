package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.components.CellMain;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.CommonWriteConfig;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.util.List;

/**
 *
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/28 11:22
 */
public interface CellStyleConfigur {

    /**
     * 配置全局样式
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。
     * 全局样式配置优先级 AutoWriteConfig内样式有关配置 > 此处配置 > 预制样式
     * @param cell  样式配置
     */
    void globalStyleConfig(CellMain cell);

    /**
     * 配置程序常用样式
     * 程序常用样式影响的范围：使用本框架提供的写入策略时由系统渲染的列或行（如：自动在结尾插入合计行、自动在第一列插入编号）
     * 部分策略不完全支持所有样式的配置 如：自动插入的编号行不支持行高的配置  合计行的样式继承上一行，只能设置行高 等
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。
     * 程序常用样式配置优先级 此处配置 > 全局样式
     * @param cell  样式配置
     */
    void commonStyleConfig(CellMain cell);

    /**
     * 配置表头样式（此处为所有表头配置样式，配置单表头样式请在Header对象内配置）
     * 渲染器渲染表头时调用
     * 表头样式配置优先级   Header对象内配置 > 此处配置 > 全局样式
     * @param cell  样式配置
     */
    void headerStyleConfig(CellMain cell);

    /**
     * 配置标题样式（标题是一个整体，此处为整个标题配置样式）
     * 渲染器渲染表头时调用
     * 标题样式配置优先级  此处配置 > 全局样式
     * @param cell  样式配置
     */
    void titleStyleConfig(CellMain cell);

    /**
     * 配置内容样式
     * 渲染内容时，每渲染一个单元格都会调用此方法
     * 内容样式配置优先级  此处配置 > 全局样式
     * @param cell  样式配置
     * @param fieldInfo 单元格与内容信息
     */
    void dataStyleConfig(CellMain cell, AbstractStyleRender.FieldInfo fieldInfo);

}

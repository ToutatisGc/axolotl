package cn.toutatis.xvoid.axolotl.excel.writer.themes.configurable;

import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlCellBorder;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlCellFont;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import org.apache.commons.beanutils.BeanUtils;

/**
 * 单元格样式配置
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/28 11:22
 */
public interface ConfigurableStyleConfig {

    /**
     * 配置全局样式
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。
     * 全局样式配置优先级 AutoWriteConfig内样式有关配置 > 此处配置 > 预制样式
     * @param cell  样式配置
     * TODO 变量名
     */
    default void globalStyleConfig(CellConfigProperty cell){}

    /**
     * 配置程序常用样式
     * 程序常用样式影响的范围：使用本框架提供的写入策略时由系统渲染的列或行（如：自动在结尾插入合计行、自动在第一列插入编号）
     * 部分策略不完全支持所有样式的配置 如：自动插入的编号行不支持行高的配置  合计行的样式继承上一行，只能设置行高 等
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。
     * 程序常用样式配置优先级 此处配置 > 全局样式
     * @param cell  样式配置
     */
    default void commonStyleConfig(CellConfigProperty cell){}

    /**
     * 配置表头样式（此处为所有表头配置样式，配置单表头样式请在Header对象内配置）
     * 渲染器渲染表头时调用
     * 表头样式配置优先级   Header对象内配置 > 此处配置 > 全局样式
     * @param cell  样式配置
     */
    default void headerStyleConfig(CellConfigProperty cell){}

    /**
     * 配置标题样式（标题是一个整体，此处为整个标题配置样式）
     * 渲染器渲染表头时调用
     * 标题样式配置优先级  此处配置 > 全局样式
     * @param cell  样式配置
     */
    default void titleStyleConfig(CellConfigProperty cell){}

    /**
     * 配置内容样式
     * 渲染内容时，每渲染一个单元格都会调用此方法
     * 内容样式配置优先级  此处配置 > 全局样式
     * @param cell  样式配置
     * @param fieldInfo 单元格与内容信息
     */
    default void dataStyleConfig(CellConfigProperty cell, AbstractStyleRender.FieldInfo fieldInfo){}

    /**
     * 导入配置
     * @param cellConfigProperty 用户配置
     * @param defaultCellProperty 默认主题配置
     * @return 导入后的主题配置
     */
    default CellPropertyHolder cloneStyleProperties(CellConfigProperty cellConfigProperty, CellPropertyHolder defaultCellProperty){
        CellPropertyHolder cellStyleProperty = copyConfigFromDefault(defaultCellProperty);
        if(cellConfigProperty.getRowHeight() != null){
            cellStyleProperty.setRowHeight(cellConfigProperty.getRowHeight());
        }
        if(cellConfigProperty.getColumnWidth() != null){
            cellStyleProperty.setColumnWidth(cellConfigProperty.getColumnWidth());
        }
        if(cellConfigProperty.getHorizontalAlignment() != null){
            cellStyleProperty.setHorizontalAlignment(cellConfigProperty.getHorizontalAlignment());
        }
        if(cellConfigProperty.getVerticalAlignment() != null){
            cellStyleProperty.setVerticalAlignment(cellConfigProperty.getVerticalAlignment());
        }
        if(cellConfigProperty.getForegroundColor() != null){
            cellStyleProperty.setForegroundColor(cellConfigProperty.getForegroundColor());
        }
        if(cellConfigProperty.getFillPatternType() != null){
            cellStyleProperty.setFillPatternType(cellConfigProperty.getFillPatternType());
        }
        AxolotlCellBorder border = cellConfigProperty.getBorder();
        if(border != null){
            if(border.getBaseBorderStyle() != null){
                cellStyleProperty.setBaseBorderStyle(border.getBaseBorderStyle());
            }
            if(border.getBaseBorderColor() != null){
                cellStyleProperty.setBaseBorderColor(border.getBaseBorderColor());
            }
            if(border.getBorderTopStyle() != null){
                cellStyleProperty.setBorderTopStyle(border.getBorderTopStyle());
            }
            if(border.getTopBorderColor() != null){
                cellStyleProperty.setTopBorderColor(border.getTopBorderColor());
            }
            if(border.getBorderBottomStyle() != null){
                cellStyleProperty.setBorderBottomStyle(border.getBorderBottomStyle());
            }
            if(border.getBottomBorderColor() != null){
                cellStyleProperty.setBottomBorderColor(border.getBottomBorderColor());
            }
            if(border.getBorderLeftStyle() != null){
                cellStyleProperty.setBorderLeftStyle(border.getBorderLeftStyle());
            }
            if(border.getLeftBorderColor() != null){
                cellStyleProperty.setLeftBorderColor(border.getLeftBorderColor());
            }
            if(border.getBorderRightStyle() != null){
                cellStyleProperty.setBorderRightStyle(border.getBorderRightStyle());
            }
            if(border.getRightBorderColor() != null){
                cellStyleProperty.setRightBorderColor(border.getRightBorderColor());
            }
        }
        AxolotlCellFont font = cellConfigProperty.getFont();
        if(font != null){
            if(font.getBold() != null){
                cellStyleProperty.setBold(font.getBold());
            }
            if(font.getFontName() != null){
                cellStyleProperty.setFontName(font.getFontName());
            }
            if(font.getFontSize() != null){
                cellStyleProperty.setFontSize(font.getFontSize());
            }
            if(font.getFontColor() != null){
                cellStyleProperty.setFontColor(font.getFontColor());
            }
            if(font.getItalic() != null){
                cellStyleProperty.setItalic(font.getItalic());
            }
            if(font.getStrikeout() != null){
                cellStyleProperty.setStrikeout(font.getStrikeout());
            }
        }
        return cellStyleProperty;
    }

    /**
     * 根据默认配置复制生成一份新配置
     * @param defaultCellProperty 默认配置
     * @return 新配置
     */
    default CellPropertyHolder copyConfigFromDefault(CellPropertyHolder defaultCellProperty){
        CellPropertyHolder cellProperty = new CellPropertyHolder();
        try {
            BeanUtils.copyProperties(cellProperty,defaultCellProperty);
        } catch (Exception e) {
            throw new AxolotlWriteException("主题配置加载失败:" + e.getMessage());
        }
        return cellProperty;
    }

}

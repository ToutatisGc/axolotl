package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.components.BaseCellProperty;
import cn.toutatis.xvoid.axolotl.excel.writer.components.CellBorder;
import cn.toutatis.xvoid.axolotl.excel.writer.components.CellFont;
import cn.toutatis.xvoid.axolotl.excel.writer.support.style.CellStyleProperty;

/**
 *
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/28 11:22
 */
public interface AxolotlCellStyleConfig {

    /**
     * 配置全局样式
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。
     * 全局样式配置优先级 AutoWriteConfig内样式有关配置 > 此处配置 > 预制样式
     * @param cell  样式配置
     */
    default void globalStyleConfig(BaseCellProperty cell){}
    default void globalStyleConfig(BaseCellProperty cell, CellStyleProperty cellStyle){
        this.cloneStyleProperties(cell,cellStyle);
    }

    /**
     * 配置程序常用样式
     * 程序常用样式影响的范围：使用本框架提供的写入策略时由系统渲染的列或行（如：自动在结尾插入合计行、自动在第一列插入编号）
     * 部分策略不完全支持所有样式的配置 如：自动插入的编号行不支持行高的配置  合计行的样式继承上一行，只能设置行高 等
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。
     * 程序常用样式配置优先级 此处配置 > 全局样式
     * @param cell  样式配置
     */
    default void commonStyleConfig(BaseCellProperty cell){}

    default void commonStyleConfig(BaseCellProperty cell, CellStyleProperty cellStyle){
        this.cloneStyleProperties(cell,cellStyle);
    }

    /**
     * 配置表头样式（此处为所有表头配置样式，配置单表头样式请在Header对象内配置）
     * 渲染器渲染表头时调用
     * 表头样式配置优先级   Header对象内配置 > 此处配置 > 全局样式
     * @param cell  样式配置
     */
    default void headerStyleConfig(BaseCellProperty cell){}

    default void headerStyleConfig(BaseCellProperty cell, CellStyleProperty cellStyle){
        this.cloneStyleProperties(cell,cellStyle);
    }

    /**
     * 配置标题样式（标题是一个整体，此处为整个标题配置样式）
     * 渲染器渲染表头时调用
     * 标题样式配置优先级  此处配置 > 全局样式
     * @param cell  样式配置
     */
    default void titleStyleConfig(BaseCellProperty cell){}

    default void titleStyleConfig(BaseCellProperty cell, CellStyleProperty cellStyle){
        this.cloneStyleProperties(cell,cellStyle);
    }


    /**
     * 配置内容样式
     * 渲染内容时，每渲染一个单元格都会调用此方法
     * 内容样式配置优先级  此处配置 > 全局样式
     * @param cell  样式配置
     * @param fieldInfo 单元格与内容信息
     */
    default void dataStyleConfig(BaseCellProperty cell, AbstractStyleRender.FieldInfo fieldInfo){}

    /**
     * 导入配置
     * @param baseCellProperty 用户配置
     * @param defaultCellStyle 主题配置
     */
    default void cloneStyleProperties(BaseCellProperty baseCellProperty, CellStyleProperty defaultCellStyle){
        if(baseCellProperty.getRowHeight() != null){
            defaultCellStyle.setRowHeight(baseCellProperty.getRowHeight());
        }
        if(baseCellProperty.getColumnWidth() != null){
            defaultCellStyle.setColumnWidth(baseCellProperty.getColumnWidth());
        }
        if(baseCellProperty.getHorizontalAlignment() != null){
            defaultCellStyle.setHorizontalAlignment(baseCellProperty.getHorizontalAlignment());
        }
        if(baseCellProperty.getVerticalAlignment() != null){
            defaultCellStyle.setVerticalAlignment(baseCellProperty.getVerticalAlignment());
        }
        if(baseCellProperty.getForegroundColor() != null){
            defaultCellStyle.setForegroundColor(baseCellProperty.getForegroundColor());
        }
        if(baseCellProperty.getFillPatternType() != null){
            defaultCellStyle.setFillPatternType(baseCellProperty.getFillPatternType());
        }
        CellBorder border = baseCellProperty.getBorder();
        if(border != null){
            if(border.getBaseBorderStyle() != null){
                defaultCellStyle.setBaseBorderStyle(border.getBaseBorderStyle());
            }
            if(border.getBaseBorderColor() != null){
                defaultCellStyle.setBaseBorderColor(border.getBaseBorderColor());
            }
            if(border.getBorderTopStyle() != null){
                defaultCellStyle.setBorderTopStyle(border.getBorderTopStyle());
            }
            if(border.getTopBorderColor() != null){
                defaultCellStyle.setTopBorderColor(border.getTopBorderColor());
            }
            if(border.getBorderBottomStyle() != null){
                defaultCellStyle.setBorderBottomStyle(border.getBorderBottomStyle());
            }
            if(border.getBottomBorderColor() != null){
                defaultCellStyle.setBottomBorderColor(border.getBottomBorderColor());
            }
            if(border.getBorderLeftStyle() != null){
                defaultCellStyle.setBorderLeftStyle(border.getBorderLeftStyle());
            }
            if(border.getLeftBorderColor() != null){
                defaultCellStyle.setLeftBorderColor(border.getLeftBorderColor());
            }
            if(border.getBorderRightStyle() != null){
                defaultCellStyle.setBorderRightStyle(border.getBorderRightStyle());
            }
            if(border.getRightBorderColor() != null){
                defaultCellStyle.setRightBorderColor(border.getRightBorderColor());
            }
        }
        CellFont font = baseCellProperty.getFont();
        if(font != null){
            if(font.getBold() != null){
                defaultCellStyle.setBold(font.getBold());
            }
            if(font.getFontName() != null){
                defaultCellStyle.setFontName(font.getFontName());
            }
            if(font.getFontSize() != null){
                defaultCellStyle.setFontSize(font.getFontSize());
            }
            if(font.getFontColor() != null){
                defaultCellStyle.setFontColor(font.getFontColor());
            }
            if(font.getItalic() != null){
                defaultCellStyle.setItalic(font.getItalic());
            }
            if(font.getStrikeout() != null){
                defaultCellStyle.setStrikeout(font.getStrikeout());
            }
        }
    }

}

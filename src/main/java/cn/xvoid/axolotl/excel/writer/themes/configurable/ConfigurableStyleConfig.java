package cn.xvoid.axolotl.excel.writer.themes.configurable;

import cn.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.xvoid.axolotl.excel.writer.components.configuration.AxolotlCellBorder;
import cn.xvoid.axolotl.excel.writer.components.configuration.AxolotlCellFont;
import cn.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import org.apache.commons.beanutils.BeanUtils;

import java.util.HashMap;
import java.util.Map;

import static cn.xvoid.axolotl.excel.writer.themes.configurable.AxolotlConfigurableTheme.DEFAULT_COLUMN_WIDTH;
import static cn.xvoid.axolotl.excel.writer.themes.configurable.AxolotlConfigurableTheme.HEADER_ROW_HEIGHT;

/**
 * 单元格样式配置
 * @author 张智凯
 * @version 1.0
 *  2024/3/28 11:22
 */
public interface ConfigurableStyleConfig {

    /**
     * 配置全局样式<p>
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。<p>
     * 全局样式配置优先级 AutoWriteConfig内样式有关配置 > 此处配置 > 预制样式<p>
     * @param cellConfig  样式配置
     */
    default void globalStyleConfig(CellConfigProperty cellConfig){}

    /**
     * 配置程序写入的单元格样式<p>
     * 使用本框架提供的写入策略时由程序写入的行或列（如：自动在结尾插入合计行、自动在第一列插入编号） 更多策略参考 ExcelWritePolicy 枚举类<p>
     * 自动在结尾插入合计行：不配置则自动取上一行单元格的样式   行高：不配置则自动取上一行的行高   列宽：不支持配置，继承上一行<p>
     * 自动在第一列插入编号：表头的'序号'单元格行高配置不生效，样式配置、列宽配置生效，
     *                    剩余的序号单元格可进行样式配置、列宽配置、行高配置(行高优先级低，不建议配置)，若无配置则使用序号单元格后第一个单元格的样式，如不存在则使用全局样式配置，
     *                    列宽可配置，可控制编号列的列宽<p>
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。<p>
     * 配置优先级 此处配置 > 全局样式<p>
     * 若要更多精细化的样式配置建议手动插入合计与编号列<p>
     * @param cellConfig  写入策略与对应单元格样式
     */
    @Deprecated
    default void commonStyleConfig(Map<ExcelWritePolicy,CellConfigProperty> cellConfig){}

    /**
     * 配置表头样式（此处为所有表头配置样式，配置单表头样式请在Header对象内配置）<p>
     * 渲染器渲染表头时调用<p>
     * 表头样式配置优先级   Header对象内配置 > 此处配置 > 全局样式<p>
     * @param cellConfig  样式配置
     */
    default void headerStyleConfig(CellConfigProperty cellConfig){}

    /**
     * 配置标题样式（标题是一个整体，此处为整个标题配置样式）<p>
     * 渲染器渲染表头时调用<p>
     * 标题样式配置优先级  此处配置 > 全局样式<p>
     * @param cellConfig  样式配置
     */
    default void titleStyleConfig(CellConfigProperty cellConfig){}

    /**
     * 配置内容样式<p>
     * 渲染内容时，每渲染一个单元格都会调用此方法<p>
     * 内容样式配置优先级  此处配置 > 全局样式<p>
     * @param cellConfig  样式配置
     * @param fieldInfo 单元格与内容信息
     */
    default void dataStyleConfig(CellConfigProperty cellConfig, AbstractStyleRender.FieldInfo fieldInfo){}

    /**
     * 导入配置
     * @param cellConfigProperty 用户配置
     * @param defaultCellProperty 默认主题配置
     * @return 导入后的主题配置
     */
    static CellPropertyHolder cloneStyleProperties(CellConfigProperty cellConfigProperty, CellPropertyHolder defaultCellProperty){
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
     * 根据默认样式配置加载程序写入的单元格样式配置
     * @param defaultCellPropHolder 默认样式配置
     * @param config 样式配置类
     * @return
     */
    static Map<ExcelWritePolicy,CellPropertyHolder> loadCommonConfigFromDefault(CellPropertyHolder defaultCellPropHolder,ConfigurableStyleConfig config){
        Map<ExcelWritePolicy, CellPropertyHolder> cellPropHolderMap = new HashMap<>();
        cellPropHolderMap.put(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,null);
        cellPropHolderMap.put(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING,null);

        Map<ExcelWritePolicy, CellConfigProperty> configPropertyMap = new HashMap<>();
        config.commonStyleConfig(configPropertyMap);
        for (ExcelWritePolicy excelWritePolicy : configPropertyMap.keySet()) {
            if(cellPropHolderMap.containsKey(excelWritePolicy)){
                CellConfigProperty configProperty = configPropertyMap.get(excelWritePolicy);
                if(configProperty != null){
                    CellPropertyHolder cellPropertyHolder = cloneStyleProperties(configProperty, defaultCellPropHolder);
                    if(cellPropertyHolder.getRowHeight() == null){
                        cellPropertyHolder.setRowHeight(HEADER_ROW_HEIGHT);
                    }
                    if(cellPropertyHolder.getColumnWidth() == null){
                        cellPropertyHolder.setColumnWidth(DEFAULT_COLUMN_WIDTH);
                    }
                    cellPropHolderMap.put(excelWritePolicy,cellPropertyHolder);
                }
            }
        }
        return cellPropHolderMap;
    }

    /**
     * 根据默认配置复制生成一份新配置
     * @param defaultCellProperty 默认配置
     * @return 新配置
     */
    static CellPropertyHolder copyConfigFromDefault(CellPropertyHolder defaultCellProperty){
        CellPropertyHolder cellProperty = new CellPropertyHolder();
        try {
            BeanUtils.copyProperties(cellProperty,defaultCellProperty);
        } catch (Exception e) {
            throw new AxolotlWriteException("主题配置加载失败:" + e.getMessage());
        }
        return cellProperty;
    }

}

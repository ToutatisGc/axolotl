package cn.toutatis.xvoid.axolotl.excel.entities.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlCellBorder;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlCellFont;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.configurable.CellConfigProperty;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.configurable.ConfigurableStyleConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Map;

/**
 * @author 张智凯
 * @version 1.0
 *  2024/3/30 3:14
 */
public class AxolotlDefaultStyleConfig implements ConfigurableStyleConfig {
    @Override
    public void globalStyleConfig(CellConfigProperty cell) {
        cell.setForegroundColor(new AxolotlColor(39,56,86));
        cell.setRowHeight((short) 550);
        AxolotlCellFont axolotlCellFont = new AxolotlCellFont();
        axolotlCellFont.setFontSize((short) 10);
        axolotlCellFont.setFontColor(IndexedColors.WHITE);
        axolotlCellFont.setFontName("微软雅黑");
        cell.setFont(axolotlCellFont);
    }

    @Override
    public void commonStyleConfig(Map<ExcelWritePolicy,CellConfigProperty> cell) {
        CellConfigProperty cellConfigProperty = new CellConfigProperty();
        AxolotlCellBorder axolotlCellBorder = new AxolotlCellBorder();
        axolotlCellBorder.setBaseBorderStyle(BorderStyle.NONE);

        axolotlCellBorder.setTopBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderTopStyle(BorderStyle.THIN);
        axolotlCellBorder.setBottomBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderBottomStyle(BorderStyle.THIN);

        cellConfigProperty.setBorder(axolotlCellBorder);
        cellConfigProperty.setForegroundColor(new AxolotlColor(255,47,47));
        cell.put(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING,cellConfigProperty);


        CellConfigProperty cellConfigProperty1 = new CellConfigProperty();
        AxolotlCellBorder axolotlCellBorder1 = new AxolotlCellBorder();
        axolotlCellBorder1.setBaseBorderStyle(BorderStyle.NONE);

        axolotlCellBorder1.setTopBorderColor(IndexedColors.BLACK);
        axolotlCellBorder1.setBorderTopStyle(BorderStyle.THIN);
        axolotlCellBorder1.setBottomBorderColor(IndexedColors.BLACK);
        axolotlCellBorder1.setBorderBottomStyle(BorderStyle.THIN);

        axolotlCellBorder1.setLeftBorderColor(IndexedColors.BLACK);
        axolotlCellBorder1.setBorderLeftStyle(BorderStyle.MEDIUM);

        cellConfigProperty1.setBorder(axolotlCellBorder1);
        cellConfigProperty1.setForegroundColor(new AxolotlColor(28,250,154));
        cell.put(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,cellConfigProperty1);
    }

    @Override
    public void headerStyleConfig(CellConfigProperty cell) {
        cell.setForegroundColor(new AxolotlColor(34,44,69));
        AxolotlCellBorder axolotlCellBorder = new AxolotlCellBorder();
        axolotlCellBorder.setBaseBorderStyle(BorderStyle.NONE);

        axolotlCellBorder.setLeftBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderLeftStyle(BorderStyle.THIN);
        axolotlCellBorder.setRightBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderRightStyle(BorderStyle.THIN);
        cell.setBorder(axolotlCellBorder);

    }

    @Override
    public void titleStyleConfig(CellConfigProperty cell) {
        cell.setRowHeight((short) 900);
        cell.setForegroundColor(new AxolotlColor( 53,80,125));
        AxolotlCellFont axolotlCellFont = new AxolotlCellFont();
        axolotlCellFont.setFontSize(StyleHelper.STANDARD_TITLE_FONT_SIZE);
        axolotlCellFont.setBold(true);
        cell.setFont(axolotlCellFont);

    }

    @Override
    public void dataStyleConfig(CellConfigProperty cell, AbstractStyleRender.FieldInfo fieldInfo) {
       // cell.setForegroundColor(new AxolotlColor(39,56,86));
        AxolotlCellBorder axolotlCellBorder = new AxolotlCellBorder();
        axolotlCellBorder.setBaseBorderStyle(BorderStyle.NONE);

        axolotlCellBorder.setTopBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderTopStyle(BorderStyle.THIN);
        axolotlCellBorder.setBottomBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderBottomStyle(BorderStyle.THIN);

        /*if(fieldInfo.getColumnIndex() == 1){
            cellBorder.setLeftBorderColor(IndexedColors.BLACK);
            cellBorder.setBorderLeftStyle(BorderStyle.MEDIUM);
        }*/
        if(fieldInfo.getColumnIndex() == 6){
            axolotlCellBorder.setRightBorderColor(IndexedColors.BLACK);
            axolotlCellBorder.setBorderRightStyle(BorderStyle.MEDIUM);
        }
        cell.setBorder(axolotlCellBorder);

    }

}

package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.CellBorder;
import cn.toutatis.xvoid.axolotl.excel.writer.components.CellFont;
import cn.toutatis.xvoid.axolotl.excel.writer.components.CellMain;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/30 3:14
 */
public class DefaultStyleConfig implements CellStyleConfigur{
    @Override
    public void globalStyleConfig(CellMain cell) {
        cell.setForegroundColor(new AxolotlColor(39,56,86));
        cell.setRowHeight((short) 550);
        CellFont cellFont = new CellFont();
        cellFont.setFontSize((short) 10);
        cellFont.setFontColor(IndexedColors.WHITE);
        cellFont.setFontName("微软雅黑");
        cell.setFont(cellFont);
    }

    @Override
    public void commonStyleConfig(CellMain cell) {
        CellBorder cellBorder = new CellBorder();
        cellBorder.setBaseBorderStyle(BorderStyle.NONE);

        cellBorder.setTopBorderColor(IndexedColors.BLACK);
        cellBorder.setBorderTopStyle(BorderStyle.THIN);
        cellBorder.setBottomBorderColor(IndexedColors.BLACK);
        cellBorder.setBorderBottomStyle(BorderStyle.THIN);

        cellBorder.setLeftBorderColor(IndexedColors.BLACK);
        cellBorder.setBorderLeftStyle(BorderStyle.MEDIUM);
        cell.setBorder(cellBorder);
    }

    @Override
    public void headerStyleConfig(CellMain cell) {
        cell.setForegroundColor(new AxolotlColor(34,44,69));
        CellBorder cellBorder = new CellBorder();
        cellBorder.setBaseBorderStyle(BorderStyle.NONE);

        cellBorder.setLeftBorderColor(IndexedColors.BLACK);
        cellBorder.setBorderLeftStyle(BorderStyle.THIN);
        cellBorder.setRightBorderColor(IndexedColors.BLACK);
        cellBorder.setBorderRightStyle(BorderStyle.THIN);
        cell.setBorder(cellBorder);

    }

    @Override
    public void titleStyleConfig(CellMain cell) {
        cell.setRowHeight((short) 900);
        cell.setForegroundColor(new AxolotlColor( 53,80,125));
        CellFont cellFont = new CellFont();
        cellFont.setFontSize(StyleHelper.STANDARD_TITLE_FONT_SIZE);
        cellFont.setBold(true);
        cell.setFont(cellFont);

    }

    @Override
    public void dataStyleConfig(CellMain cell, AbstractStyleRender.FieldInfo fieldInfo) {
       // cell.setForegroundColor(new AxolotlColor(39,56,86));
        CellBorder cellBorder = new CellBorder();
        cellBorder.setBaseBorderStyle(BorderStyle.NONE);

        cellBorder.setTopBorderColor(IndexedColors.BLACK);
        cellBorder.setBorderTopStyle(BorderStyle.THIN);
        cellBorder.setBottomBorderColor(IndexedColors.BLACK);
        cellBorder.setBorderBottomStyle(BorderStyle.THIN);

        /*if(fieldInfo.getColumnIndex() == 1){
            cellBorder.setLeftBorderColor(IndexedColors.BLACK);
            cellBorder.setBorderLeftStyle(BorderStyle.MEDIUM);
        }*/
        if(fieldInfo.getColumnIndex() == 6){
            cellBorder.setRightBorderColor(IndexedColors.BLACK);
            cellBorder.setBorderRightStyle(BorderStyle.MEDIUM);
        }
        cell.setBorder(cellBorder);

    }

}

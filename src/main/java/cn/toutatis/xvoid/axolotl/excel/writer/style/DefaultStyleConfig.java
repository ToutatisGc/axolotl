package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.CellFont;
import cn.toutatis.xvoid.axolotl.excel.writer.components.CellMain;
import org.apache.poi.ss.usermodel.IndexedColors;

/**
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/30 3:14
 */
public class DefaultStyleConfig implements CellStyleConfigur{
    @Override
    public void globalStyleConfig(CellMain cell) {
        CellFont cellFont = new CellFont();
        cellFont.setFontColor(IndexedColors.BLUE);
        cell.setFont(cellFont);
    }

    @Override
    public void commonStyleConfig(CellMain cell) {

    }

    @Override
    public void headerStyleConfig(CellMain cell) {
        cell.setForegroundColor(new AxolotlColor(247,202,142));
        CellFont cellFont = new CellFont();
        cellFont.setFontColor(IndexedColors.BLACK);
        cellFont.setBold(true);
        cellFont.setFontName("宋体");
        cellFont.setFontSize(StyleHelper.STANDARD_TITLE_FONT_SIZE);
        cell.setFont(cellFont);

    }

    @Override
    public void titleStyleConfig(CellMain cell) {
        cell.setRowHeight((short) 800);
        cell.setForegroundColor(new AxolotlColor(125,116,122));
        CellFont cellFont = new CellFont();
        cellFont.setFontColor(IndexedColors.RED);
        cellFont.setBold(true);
        cellFont.setFontName("宋体");
        cellFont.setFontSize(StyleHelper.STANDARD_TITLE_FONT_SIZE);
        cell.setFont(cellFont);
    }

    @Override
    public void dataStyleConfig(CellMain cell, AbstractStyleRender.FieldInfo fieldInfo) {


        CellFont cellFont = new CellFont();

        cellFont.setBold(true);
        cellFont.setFontName("宋体");
        cellFont.setFontSize(StyleHelper.STANDARD_TEXT_FONT_SIZE);
        cellFont.setItalic(true);
        if(x(fieldInfo.getColumnIndex(),true)){
            cellFont.setFontColor(IndexedColors.RED);
        }

        if(x(fieldInfo.getRowIndex(),false)){

        }else{
            cellFont.setFontColor(IndexedColors.GREEN);
        }
        cell.setFont(cellFont);
    }

    private boolean x(Integer start,boolean begin){
        boolean r = begin;
        for (Integer i = 0; i < start; i++) {
            if(r){
                r = false;
            }else{
                r = true;
            }
        }
        return r;
    }
}

package cn.toutatis.xvoid.axolotl.excel.writer.style;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.HashMap;
import java.util.Map;

/**
 * Excel样式助手
 * @author Toutatis_Gc
 */
public class StyleHelper {

    /**
     * 预置单元格长度
     */
    public static final Map<String,Integer> PRESET_CELL_LENGTH_MAP = new HashMap<>();

    /**
     * 默认使用字体
     */
    public static final String STANDARD_FONT_NAME = "宋体";

    /**
     * 默认起始位置
     */
    public static final int START_POSITION = 0;

    /**
     * 默认标题字体大小
     */
    public static final Short STANDARD_TITLE_FONT_SIZE = 18;

    /**
     * 默认标题行高
     */
    public static final Short STANDARD_TITLE_ROW_HEIGHT = 600;

    /**
     * 默认表头行高
     */
    public static final Short STANDARD_HEADER_ROW_HEIGHT = 350;

    /**
     * 默认文本字体大小
     */
    public static final Short STANDARD_TEXT_FONT_SIZE = 12;
    /**
     * 身份证号长度
     */
    public static final Integer SERIAL_NUMBER_LENGTH = 5120;

    /**
     * 默认文本格式化索引
     */
    public static short DATA_FORMAT_PLAIN_TEXT_INDEX = 49;


    static {
        PRESET_CELL_LENGTH_MAP.put("姓名",3500);
        PRESET_CELL_LENGTH_MAP.put("名称",5000);
        PRESET_CELL_LENGTH_MAP.put("身份证",SERIAL_NUMBER_LENGTH);
        PRESET_CELL_LENGTH_MAP.put("身份证号",SERIAL_NUMBER_LENGTH);
        PRESET_CELL_LENGTH_MAP.put("身份证号码",SERIAL_NUMBER_LENGTH);
        PRESET_CELL_LENGTH_MAP.put("性别",2000);
        PRESET_CELL_LENGTH_MAP.put("地址",12800);
    }

    /**
     * 获取预置单元格长度
     * @param cellName 单元格名称
     * @return 预置单元格长度
     */
    public static Integer getPresetCellLength(String cellName){
        return PRESET_CELL_LENGTH_MAP.getOrDefault(cellName,cellName.length()*256*8);
    }

    /**
     * 渲染合并单元格样式
     * @param sheet 工作表
     * @param cellRangeAddress 合并单元格区域
     * @param style 单元格样式
     */
    public static void renderMergeRegionStyle(Sheet sheet,CellRangeAddress cellRangeAddress, CellStyle style){
        for (int rowNum = cellRangeAddress.getFirstRow(); rowNum <= cellRangeAddress.getLastRow(); rowNum++) {
            Row currentRow = sheet.getRow(rowNum);
            if (currentRow == null) {currentRow = sheet.createRow(rowNum);}
            for (int colNum = cellRangeAddress.getFirstColumn(); colNum <= cellRangeAddress.getLastColumn(); colNum++) {
                Cell currentCell = currentRow.getCell(colNum);
                if (currentCell == null) {currentCell = currentRow.createCell(colNum);}
                currentCell.setCellStyle(style);
            }
        }
    }

    /**
     * 设置单元格水平居中,垂直居中对齐
     * @param cellStyle 单元格样式
     */
    public static void setCellStyleAlignmentCenter(CellStyle cellStyle){
        setCellStyleAlignment(cellStyle,HorizontalAlignment.CENTER,VerticalAlignment.CENTER);
    }

    /**
     * 设置单元格对齐方式
     * @param cellStyle 单元格样式
     * @param horizontalAlignment 水平对齐方式
     * @param verticalAlignment 垂直对齐方式
     */
    public static void setCellStyleAlignment(CellStyle cellStyle,HorizontalAlignment horizontalAlignment,VerticalAlignment verticalAlignment){
        cellStyle.setAlignment(horizontalAlignment);
        cellStyle.setVerticalAlignment(verticalAlignment);
    }

    /**
     * 创建工作簿字体
     * @param workbook 工作簿
     * @param fontName 字体名称
     * @param bold 是否加粗
     * @param fontSize 字体大小
     * @param fontColor 字体颜色
     */
    public static Font createWorkBookFont(Workbook workbook,String fontName,boolean bold,short fontSize,IndexedColors fontColor){
        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setBold(bold);
        font.setFontHeightInPoints(fontSize);

        font.setColor(fontColor.getIndex());
        return font;
    }

    /**
     * 创建通用的单元格样式
     * @param workbook 工作表
     * @param borderStyle 边框样式
     * @param borderColor 边框颜色
     * @return 单元格样式
     */
    public static CellStyle createCellStyle(
            Workbook workbook,
            BorderStyle borderStyle,
            IndexedColors borderColor,
            AxolotlColor foregroundColor,
            Font font
    ){
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderTop(borderStyle);
        cellStyle.setBorderRight(borderStyle);
        cellStyle.setBorderBottom(borderStyle);
        cellStyle.setBorderLeft(borderStyle);
        cellStyle.setTopBorderColor(borderColor.getIndex());
        cellStyle.setRightBorderColor(borderColor.getIndex());
        cellStyle.setBottomBorderColor(borderColor.getIndex());
        cellStyle.setLeftBorderColor(borderColor.getIndex());
        if (font != null){
            cellStyle.setFont(font);
        }
        cellStyle.setFillForegroundColor(foregroundColor.toXSSFColor());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }

    /**
     * 创建具有指定边框样式、边框颜色、前景色的单元格样式。
     *
     * @param workbook 工作簿对象，用于创建单元格样式。
     * @param borderStyle 边框样式，定义单元格边框的样式。
     * @param borderColor 边框颜色，从IndexedColors中选择颜色。
     * @param foregroundColor 前景色，定义单元格内的文字或图案颜色。
     * @return 返回配置好的单元格样式对象。
     */
    public static CellStyle createCellStyle(
            Workbook workbook,
            BorderStyle borderStyle,
            IndexedColors borderColor,
            AxolotlColor foregroundColor
    ){
        return createCellStyle(workbook, borderStyle, borderColor, foregroundColor,null);
    }

    /**
     * 创建标准单元格样式
     * @param workbook 工作簿
     * @param borderStyle 边框样式
     * @param borderColor 边框颜色
     * @param foregroundColor 背景颜色
     * @param font 字体
     * @return
     */
    public static CellStyle createStandardCellStyle(
            Workbook workbook,
            BorderStyle borderStyle,
            IndexedColors borderColor,
            AxolotlColor foregroundColor,
            Font font
    ){
        CellStyle cellStyle = createCellStyle(workbook, borderStyle, borderColor, foregroundColor, font);
        StyleHelper.setCellStyleAlignmentCenter(cellStyle);
        return cellStyle;
    }

    /**
     * 设置单元格为纯文本
     * @param cellStyle 单元格样式
     */
    public static void setCellAsPlainText(CellStyle cellStyle){
        cellStyle.setDataFormat(DATA_FORMAT_PLAIN_TEXT_INDEX);
    }

}

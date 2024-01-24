package cn.toutatis.xvoid.axolotl.excel.writer.style;

import org.apache.poi.ss.usermodel.*;

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
     * 默认标题字体大小
     */
    public static Short STANDARD_TITLE_FONT_SIZE = 18;
    /**
     * 默认文本字体大小
     */
    public static Short STANDARD_TEXT_FONT_SIZE = 12;
    /**
     * 身份证号长度
     */
    public static final Integer SERIAL_NUMBER_LENGTH = 5120;

    static {
        PRESET_CELL_LENGTH_MAP.put("姓名",3500);
        PRESET_CELL_LENGTH_MAP.put("名称",5000);
        PRESET_CELL_LENGTH_MAP.put("身份证",SERIAL_NUMBER_LENGTH);
        PRESET_CELL_LENGTH_MAP.put("身份证号",SERIAL_NUMBER_LENGTH);
        PRESET_CELL_LENGTH_MAP.put("身份证号码",SERIAL_NUMBER_LENGTH);
        PRESET_CELL_LENGTH_MAP.put("性别",1500);
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
     * 创建通用的单元格样式
     * @param workbook 工作表
     * @param borderStyle 边框样式
     * @param bold 加粗
     * @param fontName 字体名称
     * @param fontSize 字体大小
     * @param borderColor 边框颜色
     * @return 单元格样式
     */
    public static CellStyle createCommonCellStyle(
            Workbook workbook,
            BorderStyle borderStyle,
            IndexedColors borderColor,
            boolean bold,
            String fontName,
            short fontSize
    ){
        CellStyle commonCellStyle = workbook.createCellStyle();
        commonCellStyle.setBorderTop(borderStyle);
        commonCellStyle.setBorderRight(borderStyle);
        commonCellStyle.setBorderBottom(borderStyle);
        commonCellStyle.setBorderLeft(borderStyle);
        commonCellStyle.setTopBorderColor(borderColor.getIndex());
        commonCellStyle.setRightBorderColor(borderColor.getIndex());
        commonCellStyle.setBottomBorderColor(borderColor.getIndex());
        commonCellStyle.setLeftBorderColor(borderColor.getIndex());
        commonCellStyle.setAlignment(HorizontalAlignment.CENTER);
        commonCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setFontName(fontName);
        font.setBold(bold);
        font.setFontHeightInPoints(fontSize);
        commonCellStyle.setFont(font);
        return commonCellStyle;
    }

}

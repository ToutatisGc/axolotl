package cn.xvoid.axolotl.excel.writer.components.configuration;

import lombok.Getter;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * 工作簿颜色
 * @author Toutatis_Gc
 */
@Getter
public class AxolotlColor implements Color {

    private final int red;
    private final int green;
    private final int blue;

    public AxolotlColor(int red, int green, int blue) {
        this.red = red;
        this.green = green;
        this.blue = blue;
    }

    /**
     * 转换为XSSFColor
     * @return XSSFColor
     */
    public XSSFColor toXSSFColor(){
        return new XSSFColor(new byte[]{(byte)red,(byte)green,(byte)blue});
    }

    /**
     * 创建颜色
     * @param red 红色
     * @param green 绿色
     * @param blue 蓝色
     * @return 颜色
     */
    public static AxolotlColor create(int red, int green, int blue){
        return new AxolotlColor(red,green,blue);
    }

    /**
     * 创建XSSFColor颜色
     * @param red 红色
     * @param green 绿色
     * @param blue 蓝色
     * @return 颜色
     */
    public static XSSFColor createXSSFColor(int red, int green, int blue){
        return new AxolotlColor(red,green,blue).toXSSFColor();
    }

}

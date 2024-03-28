package cn.toutatis.xvoid.axolotl.excel.writer.components;

import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * 工作簿颜色
 * @author Toutatis_Gc
 */
public class AxolotlColor implements Color {

    private final int red;
    private final int green;
    private final int blue;

    public AxolotlColor(int red, int green, int blue) {
        this.red = red;
        this.green = green;
        this.blue = blue;
    }

    public int getRed() {
        return red;
    }

    public int getGreen() {
        return green;
    }

    public int getBlue() {
        return blue;
    }

    public XSSFColor toXSSFColor(){
        return new XSSFColor(new byte[]{(byte)red,(byte)green,(byte)blue});
    }

    public static AxolotlColor create(int red, int green, int blue){
        return new AxolotlColor(red,green,blue);
    }

    public static XSSFColor createXSSFColor(int red, int green, int blue){
        return new AxolotlColor(red,green,blue).toXSSFColor();
    }

}

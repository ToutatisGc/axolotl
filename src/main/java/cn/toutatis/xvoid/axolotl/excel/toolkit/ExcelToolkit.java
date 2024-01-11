package cn.toutatis.xvoid.axolotl.excel.toolkit;

/**
 * Excel工具类
 * @author Toutatis_Gc
 */
public class ExcelToolkit {

    /**
     * 获取当前读取到的行和列号的可读字符串
     * @return 当前读取到的行和列号的可读字符串
     */
    public static String getHumanReadablePosition(int rowIndex, int columnIndex) {
        char i = (char) ( 'A' + columnIndex);
        return String.format("%s", i+(String.format("%d",rowIndex + 1)));
    }

}

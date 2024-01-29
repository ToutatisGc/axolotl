package cn.toutatis.xvoid.axolotl.toolkit;

import cn.toutatis.xvoid.axolotl.Meta;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkitKt;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.slf4j.Logger;

/**
 * Excel工具类
 * @author Toutatis_Gc
 */
public class ExcelToolkit {

    private static final Logger LOGGER = LoggerToolkit.getLogger("Axolotl");

    /**
     * 获取当前读取到的行和列号的可读字符串
     * @return 当前读取到的行和列号的可读字符串
     */
    public static String getHumanReadablePosition(int rowIndex, int columnIndex) {
        char i = (char) ( 'A' + columnIndex);
        return String.format("%s", i+(String.format("%d",rowIndex + 1)));
    }

    /**
     * 判断当前行是否是空行
     * @param row 当前行
     * @return 当前行是否是空行
     */
    public static boolean blankRowCheck(Row row){
        if (row == null){
            return true;
        }
        int isAllBlank = 0;
        short lastCellNum = row.getLastCellNum();
        for (int i = 0; i < lastCellNum; i++) {
            Cell cell = row.getCell(i);
            if (cell == null || cell.getCellType() == CellType.BLANK){
                isAllBlank++;
            }else {
                return false;
            }
        }
        boolean blankRow = isAllBlank == lastCellNum;
        LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME, String.format("行[%s]数据为空",row.getRowNum()));
        return blankRow;
    }

    /**
     * 判断当前行不是空行
     * @param row 当前行
     * @return 当前行是否不是空行
     */
    public static boolean notBlankRowCheck(Row row){
        return !blankRowCheck(row);
    }

}

package cn.xvoid.axolotl.excel.writer;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.Closeable;

public interface AxolotlExcelWriter extends Closeable {

    /**
     * 刷新数据到文件中
     */
    void flush();

    /**
     * 获取工作簿
     * @return 工作簿
     */
    Workbook getWorkbook();

    /**
     * 切换工作表
     * @param sheetIndex 工作表索引
     */
    void switchSheet(int sheetIndex);

}

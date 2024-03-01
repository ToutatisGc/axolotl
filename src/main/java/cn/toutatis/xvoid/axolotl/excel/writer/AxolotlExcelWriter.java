package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.Closeable;
import java.util.List;
import java.util.Map;

public interface AxolotlExcelWriter extends Closeable {

    /**
     * 写入Excel数据
     * @param singleMap 单元格数据
     * @param circleDataList 循环引用数据
     * @return 写入结果
     * @throws AxolotlWriteException 写入异常
     */
    AxolotlWriteResult write(Map<String,?> singleMap, List<?> circleDataList) throws AxolotlWriteException;

    /**
     * 刷新数据到文件中
     */
    void flush();

    /**
     * 获取工作簿
     * @return 工作簿
     */
    SXSSFWorkbook getWorkbook();

    /**
     * 切换工作表
     * @param sheetIndex 工作表索引
     */
    void switchSheet(int sheetIndex);

}

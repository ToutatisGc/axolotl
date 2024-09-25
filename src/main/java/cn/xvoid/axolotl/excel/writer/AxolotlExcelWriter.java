package cn.xvoid.axolotl.excel.writer;

import cn.xvoid.axolotl.excel.writer.components.widgets.AxolotlImage;
import cn.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
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

    /**
     * 将图像写入指定的工作表
     *
     * @param sheetIndex 工作表的索引，指示要写入图像的工作表位置
     * @param image 要写入的图像对象，包含图像的数据和属性
     */
    void writeImage(int sheetIndex, AxolotlImage image);

    /**
     * 获取写入配置
     *
     * @return 返回一个 CommonWriteConfig 对象，包含写入操作的配置信息
     */
    CommonWriteConfig getWriteConfig();

}

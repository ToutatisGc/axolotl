package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;

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

}

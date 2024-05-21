package cn.toutatis.xvoid.axolotl.excel.reader.support.stream;

import java.util.List;

/**
 * 流数据读取任务
 * @author 张智凯
 * @version 1.0
 */
public interface ReadBatchTask<T> {

    /**
     * 执行任务
     * @param data 读取的数据
     */
    void execute(List<T> data);

}

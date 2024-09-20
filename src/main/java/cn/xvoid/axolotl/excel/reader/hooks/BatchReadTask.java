package cn.xvoid.axolotl.excel.reader.hooks;

import java.util.List;

/**
 * 定义了一个批量读取任务的接口，允许执行批量读取操作
 *
 * @param <T> 泛型参数，表示要处理的数据类型
 * @author Toutatis_Gc
 */
public interface BatchReadTask<T> {

    /**
     * 执行批量读取任务的方法
     *
     * @param data 包含要处理的数据的列表这个列表可以包含一个或多个数据项
     */
    void execute(List<T> data);

}


package cn.xvoid.axolotl.excel.reader.hooks;

/**
 * 读取进度钩子接口
 * 用于监控数据读取过程中的进度
 * @author Toutatis_Gc
 */
public interface ReadProgressHook {

    /**
     * 当读取进度更新时调用
     *
     * @param current 当前已读取的数据量
     * @param total 总数据量
     */
    void onReadProgress(int current, int total);

}

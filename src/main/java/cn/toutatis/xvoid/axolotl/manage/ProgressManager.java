package cn.toutatis.xvoid.axolotl.manage;

/**
 * 进度管理器接口，用于导入导出操作的进度管理。
 * The ProgressManager interface is designed for managing the progress of import and export operations.
 * @author Toutatis_Gc
 */
public interface ProgressManager {

    /**
     * 初始化进度管理器，设置进度ID和总记录数。
     * Initializes the progress manager with the specified progress ID and total number of records.
     *
     * @param progressId   进度ID，用于唯一标识一个进度任务。
     *                     The progress ID used to uniquely identify a progress task.
     * @param totalRecords 总记录数，表示整个操作涉及的记录总数。
     *                     The total number of records, indicating the overall number of records involved in the operation.
     */
    void init(String progressId, int totalRecords);

    /**
     * 更新进度，设置当前已处理的记录数。
     * Updates the progress by specifying the current number of processed records.
     *
     * @param progressId    进度ID，用于标识要更新的进度任务。
     *                      The progress ID identifying the progress task to be updated.
     * @param currentRecords 当前已处理的记录数，表示已经完成的记录数量。
     *                      The current number of processed records, indicating the completed number of records.
     */
    void updateProgress(String progressId, int currentRecords);

    /**
     * 获取进度。
     * @param progressId    进度ID，用于标识要更新的进度任务。
     *                      The progress ID identifying the progress task to be updated.
     * @return 当前进度。
     *         The current progress.
     */
    Double getProgress(String progressId);

    /**
     * 检查进度是否已完成。
     * Checks if the progress has been completed.
     *
     * @param progressId 进度ID，用于标识要检查的进度任务。
     *                   The progress ID identifying the progress task to be checked.
     * @return true 如果进度已完成，否则为 false。
     *         true if the progress has been completed, false otherwise.
     */
    boolean isFinished(String progressId);

}

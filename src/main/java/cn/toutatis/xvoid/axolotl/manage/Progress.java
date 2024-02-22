package cn.toutatis.xvoid.axolotl.manage;

import lombok.Getter;
import lombok.Setter;

/**
 * 进度管理器,用来获取默认实现。
 * progress manager.
 * @author Toutatis_Gc
 */
public class Progress {

    /**
     * 默认实现为内存管理进度
     */
    @Getter @Setter
    private static ProgressManager defaultProgressManager = new MemoryProgressManager();

    
    public static void init(String progressId, int totalRecords) {
        defaultProgressManager.init(progressId, totalRecords);
    }
    
    public static void updateProgress(String progressId, int currentRecords) {
        defaultProgressManager.updateProgress(progressId, currentRecords);
    }
    
    public static Double getProgress(String progressId) {
        return defaultProgressManager.getProgress(progressId);
    }
    
    public static boolean isFinished(String progressId) {
        return defaultProgressManager.isFinished(progressId);
    }
}

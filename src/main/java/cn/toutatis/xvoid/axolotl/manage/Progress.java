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

}

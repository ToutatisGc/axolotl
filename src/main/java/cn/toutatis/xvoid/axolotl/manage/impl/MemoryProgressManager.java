package cn.toutatis.xvoid.axolotl.manage.impl;

import cn.toutatis.xvoid.axolotl.exceptions.AxolotlException;
import cn.toutatis.xvoid.axolotl.manage.ProgressManager;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import org.slf4j.Logger;

import java.util.Map;
import java.util.concurrent.*;

/**
 * 进度管理器
 * 内存实现
 */
public class MemoryProgressManager implements ProgressManager {

    /**
     * 默认清理间隔(分钟)
     */
    public static final int DEFAULT_CLEAN_INTERVAL = 10;

    private final Logger LOGGER = LoggerToolkit.getLogger(MemoryProgressManager.class);


    private static final Map<String, Integer> CURRENT_RECORDS_MAP = new ConcurrentHashMap<>();

    private static final Map<String, Integer> TOTAL_PROCESS_MAP = new ConcurrentHashMap<>();

    /**
     * 已过期进程
     * 定时清理
     */
    private static final Map<String,Boolean> EXPIRED_PROCESS_MAP = new ConcurrentHashMap<>();

    /**
     * 超时进程清理器
     * 由定时线程实现
     */
    private final ScheduledExecutorService cleaner;

    public MemoryProgressManager() {
        this.cleaner = Executors.newSingleThreadScheduledExecutor();
    }

    public void initUnknownTotal() {

    }

    @Override
    public void init(String progressId, int totalRecords) {
        if (Validator.strIsBlank(progressId)){
            throw new AxolotlException("进度ID为空");
        }
        if (totalRecords <= 0){
            throw new AxolotlException("总记录数不能小于0");
        }
        boolean containsKey = CURRENT_RECORDS_MAP.containsKey(progressId);
        if (!containsKey){
            CURRENT_RECORDS_MAP.put(progressId, 0);
            TOTAL_PROCESS_MAP.put(progressId, totalRecords);
            ScheduledFuture<Integer> schedule = cleaner.schedule(() -> CURRENT_RECORDS_MAP.remove(progressId), 10, TimeUnit.MINUTES);
            cleaner.schedule(() -> TOTAL_PROCESS_MAP.remove(progressId), 10, TimeUnit.MINUTES);
        }else {
            throw new AxolotlException(
                    LoggerHelper.format("进度ID:%s已存在", progressId)
            );
        }
    }

    @Override
    public void updateProgress(String progressId, int currentRecords) {

    }

    @Override
    public Double getProgress(String progressId) {
        this.checkProgressExist(progressId);
        return (double) CURRENT_RECORDS_MAP.get(progressId) / TOTAL_PROCESS_MAP.get(progressId);
    }

    @Override
    public boolean isFinished(String progressId) {
        this.checkProgressExist(progressId);
        boolean processFinished = CURRENT_RECORDS_MAP.get(progressId) >= TOTAL_PROCESS_MAP.get(progressId);
        if (processFinished){
            CURRENT_RECORDS_MAP.remove(progressId);
            TOTAL_PROCESS_MAP.remove(progressId);
            LoggerHelper.debug(LOGGER, LoggerHelper.format("进度ID:%s已完成,将被移除", progressId));
        }
        return processFinished;
    }

    private void checkProgressExist(String progressId) {
        if (!CURRENT_RECORDS_MAP.containsKey(progressId)){
            throw new AxolotlException(
                    LoggerHelper.format("进度ID:%s不存在", progressId)
            );
        }
    }
}

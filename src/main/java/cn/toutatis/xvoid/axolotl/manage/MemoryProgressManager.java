package cn.toutatis.xvoid.axolotl.manage;

import cn.toutatis.xvoid.axolotl.exceptions.AxolotlException;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.validator.Validator;

import java.util.HashMap;
import java.util.Map;

public class MemoryProgressManager implements ProgressManager{

    private final static Map<String, Integer> CURRENT_RECORDS_MAP = new HashMap<>();

    private final static Map<String, Integer> TOTAL_PROCESS_MAP = new HashMap<>();

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
    public boolean isFinished(String progressId) {
        return false;
    }
}

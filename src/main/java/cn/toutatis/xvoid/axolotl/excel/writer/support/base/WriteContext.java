package cn.toutatis.xvoid.axolotl.excel.writer.support.base;

import cn.toutatis.xvoid.axolotl.common.AbstractContext;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.HashMap;
import java.util.Map;

/**
 * 写入上下文
 * <p>写入过程中的环境支撑</p>
 * <p>由于一个工作薄(Workbook)中可能存在多个工作表（Sheet），所以需要维护每个工作表的索引</p>
 * @author Toutatis_Gc
 */
@Data
@EqualsAndHashCode(callSuper = true)
public class WriteContext extends AbstractContext {

    /**
     * 当前写入批次
     */
    private Map<Integer,Integer> currentWrittenBatch = new HashMap<>();

    /**
     * 当前切换的sheet索引
     */
    private int switchSheetIndex = -1;

    /**
     * 是否是第一批次
     * 此方法影响读取模板
     * @return 是否第一批次写入
     */
    public boolean isFirstBatch(int sheetIndex){
        return 1 == currentWrittenBatch.get(sheetIndex);
    }

    public String getCurrentWrittenBatchAndIncrement(int sheetIndex){
        if (!currentWrittenBatch.containsKey(sheetIndex)){
            currentWrittenBatch.put(sheetIndex,0);
        }
        Integer batchNum = currentWrittenBatch.get(sheetIndex);
        currentWrittenBatch.put(sheetIndex,batchNum+1);
        return LoggerHelper.format("当前写入第[%s]批次",batchNum);
    }

}

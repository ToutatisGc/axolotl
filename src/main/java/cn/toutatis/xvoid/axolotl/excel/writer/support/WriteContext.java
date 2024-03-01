package cn.toutatis.xvoid.axolotl.excel.writer.support;

import cn.toutatis.xvoid.axolotl.common.AbstractContext;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import com.google.common.collect.HashBasedTable;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 写入上下文
 * 写入过程中的环境支撑
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
     * 同字段引用索引
     * 用于校验同一批次的字段引用去计算引用索引
     */
    private HashBasedTable<Integer,List<String>,Integer> sheetSameFields = HashBasedTable.create();

    /**
     * 单次引用索引
     * <p>单词引用数据仅能使用一次，重复写入相同字段将跳过</p>
     */
    private HashBasedTable<Integer,String,CellAddress> singleReferenceData = HashBasedTable.create();
    /**
     * 记录已使用引用索引
     */
    private HashBasedTable<Integer,String,Boolean> alreadyUsedReferenceData = HashBasedTable.create();

    /**
     * 循环引用索引
     */
    private HashBasedTable<Integer,String,CellAddress> circleReferenceData = HashBasedTable.create();

    /**
     * 当前切换的sheet索引
     */
    private int switchSheetIndex = -1;

    /**
     * 是否是模板写入
     * @return true:是模板写入 false:不是模板写入
     */
    public boolean isTemplateWrite(){
        return super.getFile() != null;
    }

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


    public void addFieldRecords(int sheetIndex,List<String> fields,int batch){
        sheetSameFields.put(sheetIndex,fields,batch);
    }

    public boolean fieldsIsInitialWriting(int sheetIndex,List<String> fields) {
        return !sheetSameFields.contains(sheetIndex,fields);
    }

}

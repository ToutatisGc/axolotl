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
     * 当前写入的行数
     */
    @Deprecated
    private int currentWrittenRow = 0;

    /**
     * 当前写入批次
     */
    private int currentWrittenBatch = 0;

    private Map<List<String>,Integer> sameFields = new HashMap<>();

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
     * 是否是模板写入
     * @return true:是模板写入 false:不是模板写入
     */
    public boolean isTemplateWrite(){
        return super.getFile() != null;
    }

    public String getCurrentWrittenBatchAndIncrement(){
        return LoggerHelper.format("当前写入第[%s]批次",++currentWrittenBatch);
    }

    public void addFieldRecords(List<String> fields,int batch){
        sameFields.put(fields,batch);
    }

    public boolean isInitialWriting(List<String> fields) {
        return !sameFields.containsKey(fields);
    }
}

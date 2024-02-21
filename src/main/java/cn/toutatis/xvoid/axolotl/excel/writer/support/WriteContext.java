package cn.toutatis.xvoid.axolotl.excel.writer.support;

import com.google.common.collect.HashBasedTable;
import lombok.Data;

import java.io.File;

/**
 * 写入上下文
 * 写入过程中的环境支撑
 * @author Toutatis_Gc
 */
@Data
public class WriteContext {

    /**
     * 模板文件
     */
    private File templateFile;

    /**
     * 当前写入的行数
     */
    @Deprecated
    private int currentWrittenRow = 0;

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
        return templateFile != null;
    }
}

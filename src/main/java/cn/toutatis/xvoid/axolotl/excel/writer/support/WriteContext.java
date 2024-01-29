package cn.toutatis.xvoid.axolotl.excel.writer.support;

import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import com.google.common.collect.HashBasedTable;
import lombok.Data;

/**
 * 写入上下文
 * 写入过程中的环境支撑
 * @author Toutatis_Gc
 */
@Data
public class WriteContext {

    /**
     * 用以读取模板的读取器
     */
    private AxolotlExcelReader<Object> templateReader;

    /**
     * 当前写入的行数
     */
    private int currentWrittenRow = 0;

    /**
     * 单次引用索引
     */
    private HashBasedTable<Integer,String,CellAddress> singleReferenceData = HashBasedTable.create();

    /**
     * 循环引用索引
     */
    private HashBasedTable<Integer,String,CellAddress> circleReferenceData = HashBasedTable.create();

    /**
     * 是否是模板写入
     * @return true:是模板写入 false:不是模板写入
     */
    public boolean isTemplateWrite(){
        return templateReader != null;
    }
}

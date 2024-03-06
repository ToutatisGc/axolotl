package cn.toutatis.xvoid.axolotl.excel.writer.support;

import com.google.common.collect.HashBasedTable;
import lombok.Data;
import lombok.EqualsAndHashCode;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
@EqualsAndHashCode(callSuper = true)
public class TemplateWriteContext extends WriteContext{

    /**
     * 已解析工作表记录
     * 存在于该记录则表示该工作表已解析过
     */
    private Map<Integer,Boolean> resolvedSheetRecord = new HashMap<>();

    /**
     * 同字段引用索引
     * 用于校验同一批次的字段引用去计算引用索引
     */
    private HashBasedTable<Integer, List<String>,Integer> sheetSameFields = HashBasedTable.create();

    /**
     * 同字段引用索引
     * 用于校验同一批次的字段引用去计算引用索引
     */
    private HashBasedTable<Integer,List<String>,List<CellAddress>> sheetNonTemplateCells = HashBasedTable.create();

    /**
     * 单次引用索引
     * <p>单词引用数据仅能使用一次，重复写入相同字段将跳过</p>
     * @see cn.toutatis.xvoid.axolotl.excel.writer.constant.TemplatePlaceholderPattern#SINGLE_REFERENCE_TEMPLATE_PLACEHOLDER
     */
    private HashBasedTable<Integer,String,CellAddress> singleReferenceData = HashBasedTable.create();

    /**
     * 循环引用索引
     * @see cn.toutatis.xvoid.axolotl.excel.writer.constant.TemplatePlaceholderPattern#CIRCLE_REFERENCE_TEMPLATE_PLACEHOLDER
     */
    private HashBasedTable<Integer,String,CellAddress> circleReferenceData = HashBasedTable.create();

    /**
     * 计算占位符
     * <p>用于计算引用数据</p>
     * @see cn.toutatis.xvoid.axolotl.excel.writer.constant.TemplatePlaceholderPattern#AGGREGATE_REFERENCE_TEMPLATE_PLACEHOLDER
     */
    private HashBasedTable<Integer,String, CellAddress> calculateReferenceData =  HashBasedTable.create();

    /**
     * 记录已使用引用索引
     */
    private HashBasedTable<Integer,String,Boolean> alreadyUsedReferenceData = HashBasedTable.create();

    /**
     * 是否是模板写入
     * @return true:是模板写入 false:不是模板写入
     */
    public boolean isTemplateWrite(){
        return super.getFile() != null;
    }


    public void addFieldRecords(int sheetIndex,List<String> fields,int batch){
        sheetSameFields.put(sheetIndex,fields,batch);
    }

    public boolean fieldsIsInitialWriting(int sheetIndex,List<String> fields) {
        return !sheetSameFields.contains(sheetIndex,fields);
    }

}

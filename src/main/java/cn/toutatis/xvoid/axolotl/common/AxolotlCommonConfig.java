package cn.toutatis.xvoid.axolotl.common;

import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import com.google.common.collect.HashBasedTable;
import lombok.Data;

import java.util.LinkedHashMap;
import java.util.Map;

@Data
public abstract class AxolotlCommonConfig {

    /**
     * sheet索引
     */
    protected int sheetIndex = 0;

    /**
     * 字典映射
     */
    protected HashBasedTable<Integer,String, Map<String,String>> dictionaryMapping = HashBasedTable.create();



    /**
     * 设置字典映射
     * @param sheetIndex sheet索引
     * @param field 字段
     * @param dict 字典
     */
    public abstract void setDict(int sheetIndex,String field,Map<String,String> dict);

    /**
     * 获取字典映射
     * @param sheetIndex sheet索引
     * @param field 字段
     * @return 字典映射
     */
    public Map<String, String> getDict(int sheetIndex, String field) {
        Map<String, String> dict = dictionaryMapping.get(sheetIndex, field);
        if(dict == null){
            return new LinkedHashMap<>();
        }
        return dict;
    }

}

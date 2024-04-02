package cn.toutatis.xvoid.axolotl.excel.writer.support.base;

import lombok.Data;

import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * 写入配置
 * @author Toutatis_Gc
 */
@Data
public class CommonWriteConfig {

    /**
     * 构造使用默认配置
     */
    public CommonWriteConfig() {
        this(true);
    }

    public CommonWriteConfig(boolean withDefaultConfig) {
        if (withDefaultConfig) {
            Map<ExcelWritePolicy, Object> defaultReadPolicies = new HashMap<>();
            for (ExcelWritePolicy policy : ExcelWritePolicy.values()) {
                if (policy.isDefaultPolicy()){
                    defaultReadPolicies.put(policy,policy.getValue());
                }
            }
            writePolicies.putAll(defaultReadPolicies);
        }
    }



    /**
     * sheet索引
     */
    private int sheetIndex = 0;

    /**
     * 写入策略
     */
    private Map<ExcelWritePolicy, Object> writePolicies = new HashMap<>();

    /**
     * 输出流
     */
    private OutputStream outputStream;


    /**
     * 添加读取策略
     */
    public void setWritePolicy(ExcelWritePolicy policy,boolean value) {
        if (policy.getType() != ExcelWritePolicy.Type.BOOLEAN){
            throw new IllegalArgumentException("读取特性不是一个布尔类型");
        }
        writePolicies.put(policy,value);
    }

    /**
     * 获取一个布尔值类型的读取策略
     */
    public boolean getWritePolicyAsBoolean(ExcelWritePolicy policy) {
        if (policy.getType() != ExcelWritePolicy.Type.BOOLEAN){
            throw new IllegalArgumentException("读取特性不是一个布尔类型");
        }
        return writePolicies.containsKey(policy) && (boolean) writePolicies.get(policy);
    }


    public void close() throws IOException {
        if (outputStream != null) {
            outputStream.close();
        }
    }

}

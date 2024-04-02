package cn.toutatis.xvoid.axolotl.excel.writer.support.base;

import com.alibaba.fastjson.JSONObject;
import lombok.Data;

/**
 * 写入结果
 * @author Toutatis_Gc
 */
@Data
public class AxolotlWriteResult {
    public AxolotlWriteResult() {
    }

    public AxolotlWriteResult(boolean write, String message) {
        this.write = write;
        this.message = message;
    }

    /**
     * 是否写入成功
     */
    private boolean write = false;

    /**
     * 写入结果信息
     */
    private String message;

    /**
     * 额外信息
     */
    private JSONObject extraInfo;

    public void setExtraInfo(String key ,Object value){
        if (this.extraInfo == null){
            this.extraInfo = new JSONObject();
        }
        this.extraInfo.put(key,value);
    }

    public Object getExtraInfo(String key){
        if (this.extraInfo == null){return null;}
        return this.extraInfo.get(key);
    }

}

package cn.toutatis.xvoid.axolotl.excel.entities.writer;

import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictKey;
import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictValue;
import lombok.Data;

@Data
public class DictTestFullEntity {

    @AxolotlDictKey
    private String dictCode;

    @AxolotlDictValue
    private String dictName;

    private String external;

    public DictTestFullEntity(String dictCode, String dictName) {
        this.dictCode = dictCode;
        this.dictName = dictName;
    }

}

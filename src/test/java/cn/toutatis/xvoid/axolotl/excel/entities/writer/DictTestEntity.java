package cn.toutatis.xvoid.axolotl.excel.entities.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlDictKey;
import lombok.Data;

@Data
public class DictTestEntity {

    @AxolotlDictKey
    private String dictCode;

    private String dictName;

    private String external;

    public DictTestEntity(String dictCode, String dictName) {
        this.dictCode = dictCode;
        this.dictName = dictName;
    }
}

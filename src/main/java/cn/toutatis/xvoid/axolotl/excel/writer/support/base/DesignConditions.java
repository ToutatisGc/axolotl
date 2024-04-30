package cn.toutatis.xvoid.axolotl.excel.writer.support.base;

import lombok.Data;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
public class DesignConditions {

    /**
     * 模板Sheet索引
     */
    private int sheetIndex;

    /**
     * 模板行高度
     */
    private Short templateLineHeight;

    /**
     * 开始写入行
     */
    private int startShiftRow;

    /**
     * 是否是Java对象
     */
    private boolean isSimplePOJO;

    /**
     * 写入字段是否为第一次写入
     */
    private boolean fieldsInitialWriting;

    /**
     * 本次写入字段Map
     */
    private Map<String,FieldInfo> writeFieldNames = new HashMap<>();

    /**
     * 写入字段列表
     */
    private List<String> writeFieldNamesList;

    /**
     * 模板行未使用字段
     */
    private Map<String, CellAddress> nonWrittenAddress;

    /**
     * 非模板单元格
     */
    private List<CellAddress> notTemplateCells;

    @Data
    public static class FieldInfo{

        private String name;

        private boolean getter = false;

        private boolean ignore = false;

        public FieldInfo() {}
        public FieldInfo(String name) {
            this.name = name;
        }
    }

}

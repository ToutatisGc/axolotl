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

    /**
     * 字段信息
     */
    @Data
    public static class FieldInfo{

        /**
         * 字段名称/Getter方法名称
         */
        private String name;

        /**
         * 是否为Getter方法
         */
        private boolean getter = false;

        /**
         * 是否忽略
         */
        private boolean ignore = false;

        /**
         * 是否存在
         */
        private boolean exist = true;

        public FieldInfo() {}
        public FieldInfo(String name) {
            this.name = name;
        }
    }

}

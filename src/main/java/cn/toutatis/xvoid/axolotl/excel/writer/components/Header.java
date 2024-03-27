package cn.toutatis.xvoid.axolotl.excel.writer.components;

import cn.toutatis.xvoid.toolkit.clazz.LambdaToolkit;
import cn.toutatis.xvoid.toolkit.clazz.XFunc;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.ArrayList;
import java.util.List;

@Data
public class Header {

    public Header(String name) {
        this.name = name;
    }

    public Header(String name, String fieldName) {
        this(name);
        this.fieldName = fieldName;
    }

    public <T,D> Header(String name, XFunc<T,D> fieldName) {
        this(name);
        this.fieldName = LambdaToolkit.getFieldName(fieldName);
    }

    public Header(String name, List<Header> childs) {
        this.name = name;
        this.childs = childs;
    }

    public Header(String name, String fieldName, List<Header> childs) {
        this.name = name;
        this.fieldName = fieldName;
        this.childs = childs;
    }

    public Header(String name, Header... childs) {
        this.name = name;
        this.childs = List.of(childs);
    }

    public Header(String name, String fieldName, Header... childs) {
        this.name = name;
        this.fieldName = fieldName;
        this.childs = List.of(childs);
    }

    /**
     * 表头标题
     */
    private String name;

    /**
     * 写入数据字段映射
     * [最底层节点生效]
     */
    private String fieldName;

    /**
     * 字段位置
     * [最底层节点生效]
     */
    private int columnPosition = -1;

    /**
     * 当前表头的子表头
     */
    private List<Header> childs = new ArrayList<>();

    /**
     * 自定义样式
     * 优先级:高
     */
    private CellStyle customCellStyle;

    /**
     * 自定义样式
     * 优先级:低
     */
    private AxolotlCellStyle axolotlCellStyle;

    /**
     * 当前表头的列宽
     * 仅最下层表头的列宽有效
     */
    private int columnWidth = -1;

    /**
     * 计算当前表头的列数
     * @return 列数
     */
    public int countOrlopCellNumber() {
        if (childs == null || childs.isEmpty()) {
            return 1; // 如果当前表头没有子表头，则返回本身
        }
        int totalCount = 0;
        for (Header subHeader : childs) {
            totalCount += subHeader.countOrlopCellNumber();
        }
        return totalCount;
    }

}

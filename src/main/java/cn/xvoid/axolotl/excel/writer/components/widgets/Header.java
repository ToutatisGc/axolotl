package cn.xvoid.axolotl.excel.writer.components.widgets;

import cn.xvoid.axolotl.excel.writer.components.configuration.AxolotlCellStyle;
import cn.xvoid.toolkit.clazz.LambdaToolkit;
import cn.xvoid.toolkit.clazz.XFunc;
import lombok.Data;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.ArrayList;
import java.util.List;

/**
 * 表头信息
 * @author Toutatis_Gc
 */
@Data
public class Header {

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
     * 是否参与计算列
     * 修改此值将加入到config中的计算列字段
     */
    private boolean participateInCalculate = false;


    public Header(String name) {
        this.name = name;
    }

    public Header(String name, boolean participateInCalculate) {
        this.name = name;
        this.participateInCalculate = participateInCalculate;
    }

    public Header(String name, String fieldName) {
        this(name);
        this.fieldName = fieldName;
    }

    public Header(String name, String fieldName, boolean participateInCalculate) {
        this.name = name;
        this.fieldName = fieldName;
        this.participateInCalculate = participateInCalculate;
    }

    public <T,D> Header(String name, XFunc<T,D> fieldName) {
        this(name);
        this.fieldName = LambdaToolkit.getFieldName(fieldName);
    }

    public <T,D> Header(String name, boolean participateInCalculate, XFunc<T,D> fieldName) {
        this(name);
        this.fieldName = LambdaToolkit.getFieldName(fieldName);
        this.participateInCalculate = participateInCalculate;
    }


    public Header(String name, List<Header> childs) {
        this.name = name;
        this.childs = childs;
    }

    public Header(String name, boolean participateInCalculate, List<Header> childs) {
        this.name = name;
        this.childs = childs;
        this.participateInCalculate = participateInCalculate;
    }

    public Header(String name, String fieldName, List<Header> childs) {
        this.name = name;
        this.fieldName = fieldName;
        this.childs = childs;
    }

    public Header(String name, String fieldName, boolean participateInCalculate, List<Header> childs) {
        this.name = name;
        this.fieldName = fieldName;
        this.childs = childs;
        this.participateInCalculate = participateInCalculate;
    }

    public Header(String name, Header... childs) {
        this.name = name;
        this.childs = List.of(childs);
    }

    public Header(String name, boolean participateInCalculate, Header... childs) {
        this.name = name;
        this.childs = List.of(childs);
        this.participateInCalculate = participateInCalculate;
    }


    public Header(String name, String fieldName, Header... childs) {
        this.name = name;
        this.fieldName = fieldName;
        this.childs = List.of(childs);
    }

    public Header(String name, String fieldName,boolean participateInCalculate, Header... childs) {
        this.name = name;
        this.fieldName = fieldName;
        this.childs = List.of(childs);
        this.participateInCalculate = participateInCalculate;
    }

    public Header name(String name) {
        this.name = name;
        return this;
    }

    public Header fieldName(String fieldName) {
        this.fieldName = fieldName;
        return this;
    }

    public Header columnPosition(int columnPosition) {
        this.columnPosition = columnPosition;
        return this;
    }

    public Header childs(Header... header) {
        this.childs = List.of(header);
        return this;
    }

    public Header childs(List<Header> childs) {
        this.childs = childs;
        return this;
    }

    public Header customCellStyle(CellStyle customCellStyle) {
        this.customCellStyle = customCellStyle;
        return this;
    }

    public Header axolotlCellStyle(AxolotlCellStyle axolotlCellStyle) {
        this.axolotlCellStyle = axolotlCellStyle;
        return this;
    }

    public Header columnWidth(int columnWidth) {
        this.columnWidth = columnWidth;
        return this;
    }

    public Header participateInCalculate(boolean participateInCalculate) {
        this.participateInCalculate = participateInCalculate;
        return this;
    }

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

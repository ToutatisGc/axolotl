package cn.toutatis.xvoid.axolotl.excel.writer.components;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class Header {

    public Header(String name) {
        this.name = name;
    }

    public Header(String name, List<Header> childs) {
        this.name = name;
        this.childs = childs;
    }

    /**
     * 表头标题
     */
    private String name;

    /**
     * 当前表头的子表头
     */
    private List<Header> childs = new ArrayList<>();

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
            return 1; // 如果当前表头没有子表头，则返回1
        }
        int totalCount = 0;
        for (Header subHeader : childs) {
            totalCount += subHeader.countOrlopCellNumber();
        }
        return totalCount;
    }

}

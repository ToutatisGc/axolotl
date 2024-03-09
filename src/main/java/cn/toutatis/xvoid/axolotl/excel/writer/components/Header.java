package cn.toutatis.xvoid.axolotl.excel.writer.components;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class Header {

    public Header(String title) {
        this.title = title;
    }

    public Header(String title, List<Header> childs) {
        this.title = title;
        this.childs = childs;
    }

    private String title;

    private List<Header> childs = new ArrayList<>();

    private int columnRange = -1;

}

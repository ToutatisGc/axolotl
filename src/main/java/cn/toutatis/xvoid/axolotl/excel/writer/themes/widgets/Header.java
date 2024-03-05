package cn.toutatis.xvoid.axolotl.excel.writer.themes.widgets;

import lombok.Data;

import java.util.List;

@Data
public class Header {

    private String title;

    private List<Header> childs;

}

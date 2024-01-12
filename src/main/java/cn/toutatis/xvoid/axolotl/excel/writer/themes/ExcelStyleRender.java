package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import org.apache.poi.xssf.streaming.SXSSFSheet;

public interface ExcelStyleRender {

    void renderHeader(SXSSFSheet sheet);
}

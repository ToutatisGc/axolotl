package cn.toutatis.xvoid.axolotl.excel.writer.style;

import org.apache.poi.xssf.streaming.SXSSFSheet;

public interface ExcelStyleRender {

    void renderHeader(SXSSFSheet sheet);
}

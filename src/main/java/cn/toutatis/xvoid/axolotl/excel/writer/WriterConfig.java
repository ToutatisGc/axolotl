package cn.toutatis.xvoid.axolotl.excel.writer;

import lombok.Data;

import java.util.List;

/**
 * 写入配置
 * @author Toutatis_Gc
 */
@Data
public class WriterConfig {

    private String title;

    private String sheetName;

    private boolean useDefaultStyle = true;

    private List<String> columnNames;


    public String getSheetName() {
        if (sheetName == null) {
            return title;
        }
        return sheetName;
    }
}

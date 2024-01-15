package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.themes.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.ExcelWriteThemes;
import lombok.Data;
import lombok.SneakyThrows;

import java.lang.reflect.InvocationTargetException;
import java.util.List;

/**
 * 写入配置
 * @author Toutatis_Gc
 */
@Data
public class WriterConfig {

    private String title;

    private String sheetName;

    private ExcelStyleRender styleRender;

    {
        try {
            styleRender = ExcelWriteThemes.$DEFAULT.getStyleRenderClass().getDeclaredConstructor().newInstance();
        } catch (InstantiationException | IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
            throw new RuntimeException(e);
        }
    }

    private List<String> columnNames;

    public void setStyleRender(ExcelStyleRender styleRender) {
        this.styleRender = styleRender;
    }

    @SneakyThrows
    public void setStyleRender(ExcelWriteThemes theme) {
        this.styleRender = theme.getStyleRenderClass().getDeclaredConstructor().newInstance();
    }

    public void setStyleRender(String themeName) {
        this.setStyleRender(ExcelWriteThemes.valueOf(themeName.toUpperCase()));
    }

    public String getSheetName() {
        if (sheetName == null) {
            return title;
        }
        return sheetName;
    }
}

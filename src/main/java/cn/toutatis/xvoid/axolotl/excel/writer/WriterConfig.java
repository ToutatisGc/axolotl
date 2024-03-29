package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.ExcelWriteThemes;
import lombok.Data;
import lombok.SneakyThrows;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

/**
 * 写入配置
 * @author Toutatis_Gc
 */
@Data
public class WriterConfig {

    /**
     * sheet索引
     */
    private int sheetIndex = 0;

    /**
     * 标题
     */
    private String title;

    /**
     * 工作表名称
     */
    private String sheetName;

    /**
     * 样式渲染器
     */
    private ExcelStyleRender styleRender;


    /**
     * 数据
     */
    private List<?> data;

    /**
     * 输出流
     */
    private OutputStream outputStream;

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

    public void close() throws IOException {
        if (outputStream != null) {
            outputStream.close();
        }
    }

}

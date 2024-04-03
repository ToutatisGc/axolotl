package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.standard.*;
import lombok.Getter;
import lombok.SneakyThrows;

/**
 * [Axolotl] Excel 写入主题
 * @author Toutatis_Gc
 */
@Getter
public enum ExcelWriteThemes {

    $DEFAULT("默认主题样式", AxolotlClassicalTheme.class),
    MIDNIGHT("午夜蓝", AxolotlMidnightTheme.class),
    SIMPLE_BLACK("经典黑", SimpleBlackTheme.class),
    ADMINISTRATION_RED("行政红", AxolotlAdministrationRedTheme.class),
    HAZE_BLUE("雾霾蓝", AxolotlHazeBlueTheme.class),
    INDUSTRIAL_ORANGE("工业橙", AxolotlIndustrialOrangeTheme.class);

    /**
     * 样式名称
     */
    private final String styleName;

    /**
     * 样式渲染器
     */
    private final Class<? extends ExcelStyleRender> styleRenderClass;

    ExcelWriteThemes(String styleName, Class<? extends ExcelStyleRender> styleRenderClass) {
        this.styleName = styleName;
        this.styleRenderClass = styleRenderClass;
    }

    /**
     * 获取样式渲染器
     * @return 样式渲染器
     */
    @SneakyThrows
    public ExcelStyleRender getRender() {
        return this.styleRenderClass.getDeclaredConstructor().newInstance();
    }

}

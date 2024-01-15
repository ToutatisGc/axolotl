package cn.toutatis.xvoid.axolotl.excel.writer.themes;

import lombok.Getter;

@Getter
public enum ExcelWriteThemes {

    $DEFAULT("默认主题样式", AxolotlTheme.class);

    private final String styleName;

    private final Class<? extends ExcelStyleRender> styleRenderClass;

    ExcelWriteThemes(String styleName, Class<? extends ExcelStyleRender> styleRenderClass) {
        this.styleName = styleName;
        this.styleRenderClass = styleRenderClass;
    }

}

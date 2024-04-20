package cn.toutatis.xvoid.axolotl.excel.writer.components.annotations;

import cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 自定义样式注解
 * 效果与AxolotlCellStyle一致
 * @author Toutatis_Gc
 * TODO 将注解转换为CellStyle
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface AxolotlCellStyle {

    /**
     * 前景色
     */
    short[] foregroundColor() default {255,255,255};

    /**
     * 填充模式
     */
    FillPatternType fillPatternType() default FillPatternType.SOLID_FOREGROUND;

    /**
     * 上边框样式
     */
    BorderStyle borderTopStyle() default BorderStyle.THIN;

    /**
     * 上边框颜色
     */
    IndexedColors topBorderColor() default IndexedColors.BLACK;

    /**
     * 下边框样式
     */
    BorderStyle borderBottomStyle() default BorderStyle.THIN;

    /**
     * 下边框颜色
     */
    IndexedColors bottomBorderColor() default IndexedColors.BLACK;

    /**
     * 左边框样式
     */
    BorderStyle borderLeftStyle() default BorderStyle.THIN;

    /**
     * 左边框颜色
     */
    IndexedColors leftBorderColor() default IndexedColors.BLACK;

    /**
     * 右边框样式
     */
    BorderStyle borderRightStyle() default BorderStyle.THIN;

    /**
     * 右边框颜色
     */
    IndexedColors rightBorderColor() default IndexedColors.BLACK;

    /**
     * 字体
     */
    String fontName() default StyleHelper.STANDARD_FONT_NAME;


    /**
     * 字体颜色
     */
    IndexedColors fontColor() default IndexedColors.BLACK;

    /**
     * 字体大小
     */
    short fontSize() default StyleHelper.STANDARD_TEXT_FONT_SIZE;

    /**
     * 字体是否加粗
     */
    boolean fontBold() default false;

    /**
     *设置文字为斜体
     */
    boolean italic() default false;

    /**
     * 使用水平删除线
     */
    boolean strikeout() default false;
}

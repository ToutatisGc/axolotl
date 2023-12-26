package cn.toutatis.xvoid.axolotl.annotations;

import cn.toutatis.xvoid.toolkit.constant.Time;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 绑定Excel单元格的属性
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface CellBindProperty {

    /**
     * 单元格序号
     */
    int cellIndex();

    /**
     * 日期格式化
     * 默认使用LocalDateTime格式
     */
    String dateFormat() default Time.SIMPLE_DATE_FORMAT_REGEX;
}

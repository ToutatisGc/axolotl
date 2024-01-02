package cn.toutatis.xvoid.axolotl.excel.annotations;

import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.adapters.AutoAdapter;
import cn.toutatis.xvoid.toolkit.constant.Time;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 标注在字段上，指定该字段在excel中的位置
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface SpecifyPositionBind {

    /**
     * 指定单元格位置，如A1,B2等
     */
    String value();

    String format() default Time.SIMPLE_DATE_FORMAT_REGEX;

    /**
     * 指定单元格位置的适配器，默认使用默认适配器
     */
    Class<? extends DataCastAdapter<?>> adapter() default AutoAdapter.class;
}

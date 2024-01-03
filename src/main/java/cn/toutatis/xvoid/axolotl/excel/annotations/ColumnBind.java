package cn.toutatis.xvoid.axolotl.excel.annotations;

import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.adapters.AutoAdapter;
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
public @interface ColumnBind {

    /**
     * 单元格序号
     */
    int columnIndex();

    /**
     * 日期格式化
     * 默认使用LocalDateTime格式
     */
    String format() default Time.SIMPLE_DATE_FORMAT_REGEX;

    /**
     * 自定义适配器
     * @see cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter
     */
    Class<? extends DataCastAdapter<?>> adapter() default AutoAdapter.class;

}

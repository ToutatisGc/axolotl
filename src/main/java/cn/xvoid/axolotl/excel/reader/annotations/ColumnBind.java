package cn.xvoid.axolotl.excel.reader.annotations;

import cn.xvoid.axolotl.excel.reader.support.DataCastAdapter;
import cn.xvoid.axolotl.excel.reader.support.adapters.AutoAdapter;
import cn.xvoid.toolkit.constant.Time;

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
     * 表头名称
     * @return 表头名称
     */
    String[] headerName() default {};
    /**
     * 单元格序号
     */
    int columnIndex() default -1;

    /**
     * 有多个相同的表头可使用此索引指定位置
     */
    int sameHeaderIdx() default -1;

    /**
     * 数据格式化
     * 默认使用LocalDateTime格式
     */
    String format() default Time.YMD_HORIZONTAL_FORMAT_REGEX;

    /**
     * 自定义适配器
     * @see DataCastAdapter
     */
    Class<? extends DataCastAdapter<?>> adapter() default AutoAdapter.class;

}

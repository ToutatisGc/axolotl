package cn.toutatis.xvoid.axolotl.excel.writer.components.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 简单表头
 * 当使用实体写入时，该注解将作为表头显示
 * 若写入时指定了headers，该注解将被忽略
 * @author Toutatis_Gc
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface SheetHeader {

    /**
     * 表头名称
     */
    String name();

    /**
     * 列宽
     * -1表示使用默认值
     */
    int width() default -1;

}

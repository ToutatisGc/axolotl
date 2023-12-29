package cn.toutatis.xvoid.axolotl.excel.annotations;

import kotlin.annotation.MustBeDocumented;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 表索引指定注解
 */
@WorkSheet
@MustBeDocumented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface IndexWorkSheet {

    /**
     * 指定表索引
     */
    int sheetIndex() default 0;

}

package cn.toutatis.xvoid.axolotl.annotations;

import kotlin.annotation.MustBeDocumented;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

/**
 * 表名称指定注解
 */
@WorkSheet
@MustBeDocumented
@Target(ElementType.TYPE)
public @interface NamingWorkSheet {

    /**
     * 指定表名称
     */
    String sheetName() default "";


}

package cn.toutatis.xvoid.axolotl.annotations;

import kotlin.annotation.MustBeDocumented;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

/**
 * Excel注解
 */
@MustBeDocumented
@Target(ElementType.TYPE)
public @interface WorkSheet {

    String sheetName() default "";

    int sheetIndex() default 0;

}

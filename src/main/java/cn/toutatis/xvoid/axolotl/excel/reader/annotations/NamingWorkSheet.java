package cn.toutatis.xvoid.axolotl.excel.reader.annotations;

import kotlin.annotation.MustBeDocumented;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 表名称指定注解
 */
@WorkSheet
@MustBeDocumented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface NamingWorkSheet {

    /**
     * 指定表名称
     */
    String sheetName();


}

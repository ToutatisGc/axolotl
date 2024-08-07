package cn.xvoid.axolotl.excel.reader.annotations;

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

    /**
     * 读取起始偏移行
     */
    int readRowOffset() default 0;

    /**
     * 工作表列有效范围
     */
    int[] sheetColumnEffectiveRange() default {0,-1};

}

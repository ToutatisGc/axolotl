package cn.toutatis.xvoid.axolotl.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 标注在字段上，指定该字段在excel中的位置
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface SpecifyCellPosition {

    /**
     * 指定单元格位置，如A1,B2等
     */
    String value();
}

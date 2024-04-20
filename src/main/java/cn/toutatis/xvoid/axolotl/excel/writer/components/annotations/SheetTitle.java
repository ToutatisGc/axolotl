package cn.toutatis.xvoid.axolotl.excel.writer.components.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel写入标题信息
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface SheetTitle {

    /**
     * 自动生成时指定标题
     * @return 表头
     */
    String value();

    /**
     * 指定Excel工作表名称
     * 不指定时使用value()
     * @return 工作表名称
     */
    String sheetName() default "";

}

package cn.xvoid.axolotl.excel.writer.components.annotations;

import java.lang.annotation.*;

/**
 * @author 张智凯
 * @version 1.0
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface AxolotlWriterGetter {

    /**
     * 未配置使用字段Getter
     */
    String value();

}

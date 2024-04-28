package cn.toutatis.xvoid.axolotl.excel.reader.annotations;

import java.lang.annotation.*;

/**
 * 使用 TODO 特性
 * @author 张智凯
 * @since 1.0.15
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface AxolotlReaderSetter {

    /**
     * 未配置使用字段同名Setter
     */
    String value();

}

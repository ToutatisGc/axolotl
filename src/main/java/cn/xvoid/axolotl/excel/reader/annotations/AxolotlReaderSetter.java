package cn.xvoid.axolotl.excel.reader.annotations;

import java.lang.annotation.*;

/**
 * 用于标注在字段上，以提供特定的setter方法配置。
 * 允许在读取器运行时通过反射获取到该注解，并执行相应的逻辑。
 * @author 张智凯
 * @since 1.0.15
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface AxolotlReaderSetter {

    /**
     * 该字段用于指定是否使用字段同名的setter方法。
     * 如果配置了具体的值（非空），则表示使用该值作为setter方法名；
     * 如果为空，则表示使用字段同名的setter方法。
     *
     * @return String 表示setter方法的名字。如果不配置，默认为字段同名。
     */
    String value();

}

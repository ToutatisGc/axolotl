package cn.toutatis.xvoid.axolotl.common.annotations;

import java.lang.annotation.*;

/**
 * 字典映射策略 配置
 * @author Toutatis_Gc
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface AxolotlDictMappingPolicy {

    /**
     * 字典未匹配到的默认值
     */
    String defaultValue() default "";

    /**
     * 字典映射策略
     */
    DictMappingPolicy mappingPolicy() default DictMappingPolicy.KEEP_ORIGIN;

}

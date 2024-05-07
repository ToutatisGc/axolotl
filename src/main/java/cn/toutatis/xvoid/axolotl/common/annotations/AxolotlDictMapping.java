package cn.toutatis.xvoid.axolotl.common.annotations;

import java.lang.annotation.*;

/**
 * 字典映射
 * @author Toutatis_Gc
 * @since 1.0.15
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD,ElementType.METHOD})
public @interface AxolotlDictMapping {

    /**
     * 映射字段名
     */
    String value() default "";

    /**
     * 标记字段是否被映射
     */
    boolean isUsage() default true;

    /**
     * 字典映射生效的工作表
     */
    int[] effectSheetIndex() default {};

    /**
     * 静态字典
     * <p>适合明确的字典映射</p>
     * <p>格式为{"TEST_01","字典01","TEST_02":"字典02"}</p>
     * <p>如果为动态查询,应当使用config配置字典映射</p>
     */
    String[] staticDict() default {};

    /**
     * <p>是否使用手动配置优先级</p>
     * <p>手动配置优先级高于静态字典</p>
     */
    boolean useManualConfigPriority() default true;

    /**
     * 字典未匹配到的默认值
     * 只配置一个值即可，忽略其他值类型
     */
    String[] defaultValue() default {};

    /**
     * 字典映射策略
     */
    DictMappingPolicy mappingPolicy() default DictMappingPolicy.KEEP_ORIGIN;

}

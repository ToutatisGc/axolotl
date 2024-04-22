package cn.toutatis.xvoid.axolotl.common.annotations;

import java.lang.annotation.*;

/**
 * TODO 字典映射功能
 * 字典映射
 * @author Toutatis_Gc
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
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

    boolean useManualConfigPriority() default true;

}

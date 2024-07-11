package cn.toutatis.xvoid.axolotl.common.annotations;

import java.lang.annotation.*;

/**
 * 手动指定字典是否翻转
 * @author Toutatis_Gc
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD,ElementType.METHOD})
public @interface AxolotlDictOverTurn {

    Class<?>[] value() default {};
}

package cn.toutatis.xvoid.axolotl.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

/**
 * 保留原有数据，不做任何修改
 */
@Target(ElementType.FIELD)
public @interface KeepIntact {
}

package cn.toutatis.xvoid.axolotl.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

/**
 * 绑定Excel单元格的属性
 */
@Target(ElementType.FIELD)
public @interface CellBindProperty {

    int cellIndex();

}

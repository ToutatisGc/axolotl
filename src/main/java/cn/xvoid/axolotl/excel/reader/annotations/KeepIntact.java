package cn.xvoid.axolotl.excel.reader.annotations;

import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 保留原有数据，不做任何修改
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD,ElementType.TYPE})
public @interface KeepIntact {

    /**
     * 排除的读取特性
     * @return 排除的读取特性
     */
    ExcelReadPolicy[] excludePolicies();
}

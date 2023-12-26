package cn.toutatis.xvoid.axolotl.annotations;

import cn.toutatis.xvoid.axolotl.constant.ReadExcelFeature;

import java.lang.annotation.ElementType;
import java.lang.annotation.Target;

/**
 * 保留原有数据，不做任何修改
 */
@Target({ElementType.FIELD,ElementType.TYPE})
public @interface KeepIntact {

    /**
     * 排除的读取特性
     * @return 排除的读取特性
     */
    ReadExcelFeature[] excludeFeatures();
}

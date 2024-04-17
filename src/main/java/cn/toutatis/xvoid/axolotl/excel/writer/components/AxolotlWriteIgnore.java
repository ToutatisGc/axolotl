package cn.toutatis.xvoid.axolotl.excel.writer.components;

import java.lang.annotation.*;

/**
 * TODO 忽略写入
 * 写入器使用SIMPLE_USE_GETTER_METHOD特性时所忽略的getter方法或字段
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD,ElementType.METHOD})
public @interface AxolotlWriteIgnore {}
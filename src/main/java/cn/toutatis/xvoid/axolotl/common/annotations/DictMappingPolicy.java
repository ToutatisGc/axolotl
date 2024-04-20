package cn.toutatis.xvoid.axolotl.common.annotations;

/**
 * 字典映射策略
 * 字典匹配到值则使用字典值，否则使用策略决定赋值
 * @author 张智凯
 * @version 1.0
 */
public enum DictMappingPolicy {

    /**
     * 使用注解中配置的默认值
     */
    USE_DEFAULT,

    /**
     * 保持字段值原值
     */
    KEEP_ORIGIN,

    /**
     * 设置为空
     */
    NULL_VALUE

}

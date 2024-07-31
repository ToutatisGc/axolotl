package cn.xvoid.axolotl.toolkit;


import jakarta.validation.*;

import java.util.Set;

/**
 * 实体验证器工具类，负责对实体对象进行验证。
 * 采用单例模式实现，确保整个应用中只有一个实例存在。
 * @author Toutatis_Gc
 * @since 1.0.16
 */
public class EntityValidator {

    /**
     * 静态volatile变量，用于实现单例模式。
     */
    private static volatile EntityValidator instance;

    /**
     * 用于实体验证的Hibernate Validator实例。
     */
    private static Validator validator;

    /**
     * 私有构造方法，防止外部实例化对象。
     */
    private EntityValidator() {}

    /**
     * 获取EntityValidator的单例实例。
     * 如果尚未初始化，则通过双重检查锁定进行初始化。
     *
     * @return EntityValidator的单例实例。
     */
    public static EntityValidator INSTANCE() {
        if (instance == null) {
            synchronized (EntityValidator.class) {
                if (instance == null) {
                    instance = new EntityValidator();
                    try (ValidatorFactory validatorFactory = Validation.buildDefaultValidatorFactory()) {
                        validator = validatorFactory.getValidator();
                    }
                }
            }
        }
        return instance;
    }

    /**
     * 对给定的实体对象进行验证。
     * 如果验证失败，即存在约束违规情况，则抛出ValidationException异常。
     *
     * @param entity 待验证的实体对象。
     * @param groups 验证的组，用于指定特定的验证规则。
     * @throws ValidationException 如果验证失败，则抛出此异常。
     */
    public <T> void validate(T entity, Class<?>... groups) {
        Set<ConstraintViolation<T>> validate = validator.validate(entity, groups);
        if (!validate.isEmpty()) {
            throw new ValidationException(validate.iterator().next().getMessage());
        }
    }

}


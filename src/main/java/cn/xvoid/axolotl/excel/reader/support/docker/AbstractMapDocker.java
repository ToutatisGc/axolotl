package cn.xvoid.axolotl.excel.reader.support.docker;

import lombok.Data;

/**
 * 抽象类AbstractMapDocker为MapDocker接口提供了一个基本的实现框架。
 * 该类用于处理与映射相关的docker操作，允许子类具体实现这些操作。
 *
 * @author Toutatis_Gc
 * @param <T> 泛型参数T，表示该类中操作的对象类型。
 */
@Data
public abstract class AbstractMapDocker<T> implements MapDocker<T> {

    /**
     * 控制是否显示null值的标志位。
     * 当nullDisplay为true时，表示在某些操作中需要显示null值；
     * 当nullDisplay为false时，则在操作中不需要显示null值。
     */
    private Boolean nullDisplay;

}

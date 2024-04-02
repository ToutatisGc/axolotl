package cn.toutatis.xvoid.axolotl.excel.writer.components;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;

/**
 * 下拉框支持
 * TODO 校验等方法
 * @author Toutatis_Gc
 * @param <T> 选项泛型
 */
@Data
@AllArgsConstructor
public class AxolotlSelectBox<T> {

    /**
     * 下拉框值
     */
    private T value;

    /**
     * 下拉框选项
     */
    private List<T> options;

}

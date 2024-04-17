package cn.toutatis.xvoid.axolotl.excel.writer.components;

import cn.toutatis.xvoid.axolotl.excel.writer.support.inverters.DataInverter;
import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * 下拉框支持
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


    /**
     * 将属性转换为字符串类型
     * @return 转换后的对象
     */
    public AxolotlSelectBox<String> convertPropertiesToString(DataInverter<?> dataInverter){
        String value = null;
        List<String> options = new ArrayList<>();
        if(this.value != null){
            value = dataInverter.convert(this.value).toString();
        }
        if(this.options != null){
            options = this.options.stream().filter(Objects::nonNull).map(o -> dataInverter.convert(o).toString()).collect(Collectors.toList());
        }

        AxolotlSelectBox<String> selectBox = new AxolotlSelectBox<>(value, options);
        if((!selectBox.getOptions().isEmpty()) && (!selectBox.getOptions().contains(value))){
            selectBox.setValue(null);
        }
        return selectBox;
    }

}

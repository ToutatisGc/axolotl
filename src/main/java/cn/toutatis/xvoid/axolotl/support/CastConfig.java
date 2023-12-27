package cn.toutatis.xvoid.axolotl.support;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class CastConfig<T> {

    private Class<T> castType;

    private String dataFormat;

}

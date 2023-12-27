package cn.toutatis.xvoid.axolotl.support;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class CastContext<T> {

    private Class<T> castType;

    private String dataFormat;

}

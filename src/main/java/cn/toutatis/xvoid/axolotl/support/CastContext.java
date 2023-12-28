package cn.toutatis.xvoid.axolotl.support;

import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * 用于存放cast相关的上下文信息
 * @param <FT> 需要转换的字段类型
 * @author Toutatis_Gc
 */
@Data
@AllArgsConstructor
public class CastContext<FT> {

    private Class<FT> castType;

    private String dataFormat;

}

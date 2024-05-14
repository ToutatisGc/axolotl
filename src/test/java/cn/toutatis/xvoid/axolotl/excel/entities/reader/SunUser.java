package cn.toutatis.xvoid.axolotl.excel.entities.reader;

import cn.toutatis.xvoid.axolotl.excel.reader.annotations.AxolotlReaderSetter;
import lombok.Data;

/**
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/1 16:03
 */
@Data
public class SunUser {

    @AxolotlReaderSetter("")
    private String username;

    private String personName;

    private String cardNumberDec;

    private String phone;

    private String approvalStatus;

    private String bankName;

    private String source;

    private String createTime;
}

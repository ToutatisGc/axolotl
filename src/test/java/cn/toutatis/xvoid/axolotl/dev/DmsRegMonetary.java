package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.excel.annotations.SpecifyPositionBind;
import lombok.Data;

import java.io.Serial;
import java.io.Serializable;

@Data
public class DmsRegMonetary implements Serializable {

    @Serial
    private static final long serialVersionUID = 1L;

    private String formId;
    /**
     * 现金账面余额
     */
    @SpecifyPositionBind("E7")
    private String monetaryResourcesCashBookBalance;
    /**
     * 加：已收未入账
     */
    @SpecifyPositionBind("E8")
    private String monetaryResourcesReceivedNotRecorded;
    /**
     * 减：已支未入账
     */
    @SpecifyPositionBind("E9")
    private String monetaryResourcesPaidNotRecorded;
    /**
     * 银行存款账面余额
     */
    @SpecifyPositionBind("I11")
    private String monetaryResourcesBankBookBalance;

}

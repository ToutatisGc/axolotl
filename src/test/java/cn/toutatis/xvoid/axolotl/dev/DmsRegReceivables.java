package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.util.Date;


@Data
public class DmsRegReceivables{

    private String formId;
    
    /**
     * 债务人
     */
    @ColumnBind(columnIndex = 1)
    private String receivablesDebtor;
    /**
     * 形成原因
     */
    @ColumnBind(columnIndex = 2)
    private String receivablesCauses;
    /**
     * 到期时间
     */
    @ColumnBind(columnIndex = 3)
    private String receivablesExpirationTimeString;

    @ColumnBind(columnIndex = 3)
    private LocalDateTime receivablesExpirationLocalDateTime;

    @ColumnBind(columnIndex = 3)
    private Date receivablesExpirationDate;
    /**
     * 审批人
     */
    @ColumnBind(columnIndex = 4)
    private String receivablesApprover;
    /**
     * 账面数
     */
    @ColumnBind(columnIndex = 5)
    private String receivablesPapernumber;
    /**
     * 清查核实增加
     */
    @ColumnBind(columnIndex = 6)
    private String receivablesCheckAdd;
    /**
     * 清查核实减少
     */
    @ColumnBind(columnIndex = 7)
    private String receivablesCheckReduce;
    /**
     * 核实数
     */
    @ColumnBind(columnIndex = 8)
    private String receivablesVerify;

    @ColumnBind(columnIndex = 8)
    private BigDecimal receivablesVerifyBigDecimal;
    /**
     * 备注
     */
    @ColumnBind(columnIndex = 9)
    private String receivablesRemark;



}

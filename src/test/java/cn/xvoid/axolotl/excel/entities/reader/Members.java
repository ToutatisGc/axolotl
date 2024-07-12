package cn.xvoid.axolotl.excel.entities.reader;

import cn.xvoid.axolotl.excel.reader.annotations.ColumnBind;

import cn.xvoid.axolotl.excel.reader.annotations.IndexWorkSheet;
import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Objects;

/**
 * 集体经济成员信息管理表
 */
@Data
@IndexWorkSheet(readRowOffset = 5)
public class Members {

    private String kid; // 主键

    @ColumnBind(columnIndex = 6)
    private String membersFamilyRelationship; // 成员与户代表关系

    @ColumnBind(columnIndex = 12)
    private String phone; // 联系电话

    private String bankCode; // 机构号

    @ColumnBind(columnIndex = 2)
    private String membersName; // 成员姓名

    @ColumnBind(columnIndex = 1)
    private String membersFamilyCode; // 成员户编码

    @ColumnBind(columnIndex = 3)
    private String sex; // 性别

    @ColumnBind(columnIndex = 5)
    private String cardNumber; // 身份证号

    @ColumnBind(columnIndex = 4)
    private String membersNation; // 民族

    private String memberStatus; // 成员状态（字典）

    @ColumnBind(columnIndex = 7)
    private LocalDate thatTime; // 确认时间

    @ColumnBind(columnIndex = 11)
    private String address; // 住址

    @ColumnBind(columnIndex = 8)
    private String ifHave; // 是否持有农村集体经营性资产收益分配权份额（股份）(0:否 1:是)

    @ColumnBind(columnIndex = 9)
    private BigDecimal contractedArea; // 承包地确权总面积

    @ColumnBind(columnIndex = 10)
    private BigDecimal homesteadLandArea; // 宅基地土地使用权面积

    private String openingBankName; // 开户名

    private String openingBank; // 开户行

    private String creditCardNumber; // 银行卡号

    private String shareRightNumber; // 股权证书编号

    private String shareRightAttachments; // 股权证书文件

    @ColumnBind(columnIndex = 13)
    private String remark; // 备注

    /**
     * 新增原因
     */
    private String reason;

    /**
     * 变更原因
     */
    private String updateReason;

    /**
     * 加密身份证号
     */
    private String encryptCardNumber;

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Members members = (Members) o;
        return Objects.equals(cardNumber, members.cardNumber);
    }

    @Override
    public int hashCode() {
        return Objects.hash(cardNumber);
    }
}


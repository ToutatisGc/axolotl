package cn.xvoid.axolotl.excel.entities.reader;

import cn.xvoid.axolotl.excel.reader.annotations.ColumnBind;

import cn.xvoid.axolotl.excel.reader.annotations.IndexWorkSheet;
import lombok.Data;
import lombok.experimental.Accessors;

import java.io.Serializable;
import java.time.LocalDateTime;

/**
 * @description: 网格员实体类
 * @date: 2024/7/15 11:56
 * @author: liuxiaoxia
 */
@Data
@IndexWorkSheet(readRowOffset = 1)
public class DisGridInspectorEntity implements Serializable {

    /** 主键 */
    private String kid;
    /** 乡镇 */
    @ColumnBind(columnIndex = 0)
    private String town;
    /** 村 */
    @ColumnBind(columnIndex = 1)
    private String village;
    /** 网格编号 */
    @ColumnBind(columnIndex = 2)
    private String gridNumbering;
    /** 姓名 */
    @ColumnBind(columnIndex = 3)
    private String name;
    /** 电话 */
    @ColumnBind(columnIndex = 4)
    private String phone;
    /** 性别 */
    @ColumnBind(columnIndex = 5)
    private String sex;
    /** 身份证加密 */
    private String identity;
    /** 身份 */
    @ColumnBind(columnIndex = 6)
    private String identityCard;
    /** 政治面貌 */
    @ColumnBind(columnIndex = 7)
    private String politicsStatus;
    /** 学历 */
    @ColumnBind(columnIndex = 8)
    private String educationBackground;
    /** 户数 */
    @ColumnBind(columnIndex = 9)
    private Integer households;
    /** 人数 */
    @ColumnBind(columnIndex = 10)
    private Integer personNumber;
    /** 总户数 */
    @ColumnBind(columnIndex = 11)
    private Integer totalHouseholds;
    /** 总人数 */
    @ColumnBind(columnIndex = 12)
    private Integer totalPerson;
    /** 四至范围 */
    @ColumnBind(columnIndex = 13)
    private String fourRange;
    /** 用户id */
    private String userKid;
    /** 机构号 */
    private String bankCode;
}

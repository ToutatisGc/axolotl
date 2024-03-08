package cn.toutatis.xvoid.axolotl.excel.entities.writer;

import lombok.Data;

import java.io.Serializable;

/**
 * 上报检测
 */
@Data
public class MpOrgDataIssueNew implements Serializable {

    private String orgNo = "014";

    private String bankName = "山西省";

    private String dataIssue = "2024-02";

    private String shouldCnt = "1";

    private String overCnt = "2";

    private String notCnt = "3";

    private String scheduleRate = "4";

    private String hasChildren = "1";

    private String upCode;

    private String bankLevel;

    private String dataCnt = "5";

    private String vlgCnt = "6";

}

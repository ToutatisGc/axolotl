package cn.xvoid.axolotl.excel.entities.writer;

import cn.xvoid.axolotl.excel.writer.components.annotations.AxolotlWriteIgnore;
import cn.xvoid.axolotl.excel.writer.components.annotations.AxolotlWriterGetter;
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

    @AxolotlWriterGetter("getNotCnt")
    private String overCnt = "2";

    @AxolotlWriteIgnore
    private String notCnt = "3";

    private String scheduleRate = "4";

    private String hasChildren = "1";

    private String upCode;

    private String bankLevel;

    private String dataCnt = "5";

    private String vlgCnt = "6";

    private String regionStatus = "ST_001";


    public String getOrgNo() {
        return orgNo;
    }

    public String getBankName() {
        return bankName;
    }

    public String getDataIssue() {
        return dataIssue;
    }

    public String getShouldCnt() {
        System.err.println("使用GetterShouldCnt");
        return shouldCnt;
    }

    public String getOverCnt() {
        return overCnt;
    }

    public String getNotCnt() {
        return notCnt;
    }

    public String getScheduleRate() {
        return scheduleRate;
    }
}

package cn.toutatis.xvoid.axolotl.support;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;

/**
 * 工作簿元信息
 * @author Toutatis_Gc
 */
public class WorkBookMetaInfo extends AbstractMetaInfo{

    private Workbook workbook;

    public WorkBookMetaInfo(File file, DetectResult detectResult) {
        this.setFile(file);
        this.setMimeType(detectResult.getCatchMimeType());
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }
}

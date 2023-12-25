package cn.toutatis.xvoid.axolotl.support;

import lombok.Getter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;

/**
 * 工作簿元信息
 * @author Toutatis_Gc
 */
public class WorkBookMetaInfo extends AbstractMetaInfo{

    @Getter
    private Workbook workbook;

    private FormulaEvaluator formulaEvaluator;

    public WorkBookMetaInfo(File file, DetectResult detectResult) {
        this.setFile(file);
        this.setMimeType(detectResult.getCatchMimeType());
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public FormulaEvaluator getFormulaEvaluator() {
        if (formulaEvaluator == null) {
            formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        }
        return formulaEvaluator;
    }
}

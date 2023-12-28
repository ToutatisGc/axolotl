package cn.toutatis.xvoid.axolotl.support;

import lombok.Getter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 工作簿元信息
 * @author Toutatis_Gc
 */
public class WorkBookMetaInfo extends AbstractMetaInfo{

    @Getter
    private Workbook workbook;

    private FormulaEvaluator formulaEvaluator;

    private final Map<Integer, List<Row>> sheetData = new HashMap<>();

    /**
     * 当前读取到的行和列号
     * -1表示未读取到行号和列号
     * 这两种属性用于在读取数据时，获取当前读取到的行和列号，以提示错误信息和定位读取位置
     */
    private int currentReadRowIndex = -1;
    private int currentReadColumnIndex = -1;

    public WorkBookMetaInfo(File file, DetectResult detectResult) {
        this.setFile(file);
        this.setMimeType(detectResult.getCatchMimeType());
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * 获取工作簿公式计算器
     * @return 公式计算器
     */
    public FormulaEvaluator getFormulaEvaluator() {
        if (formulaEvaluator == null) {
            formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
        }
        return formulaEvaluator;
    }

    public void setSheetData(int sheetIndex,List<Row> sheetData) {
        this.sheetData.put(sheetIndex, sheetData);
    }

    /**
     * 获取工作簿指定sheet的数据
     * @param sheetIndex sheet序号
     * @return sheet数据
     */
    public List<Row> getSheetData(int sheetIndex) {
        return sheetData.getOrDefault(sheetIndex, null);
    }

    /**
     * 判断指定sheet的数据是否为空
     * @param sheetIndex sheet序号
     * @return sheet数据是否为空
     */
    public boolean isSheetDataEmpty(int sheetIndex) {
        return !sheetData.containsKey(sheetIndex);
    }
}

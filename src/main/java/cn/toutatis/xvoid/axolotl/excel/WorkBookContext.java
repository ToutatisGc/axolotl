package cn.toutatis.xvoid.axolotl.excel;

import cn.toutatis.xvoid.axolotl.excel.reader.support.AbstractContext;
import cn.toutatis.xvoid.axolotl.excel.toolkit.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.excel.toolkit.ExcelToolkit;
import lombok.Getter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;

/**
 * 工作簿元信息
 * @author Toutatis_Gc
 */
public class WorkBookContext extends AbstractContext {

    /**
     * 由文件加载而来的工作簿文件信息
     */
    @Getter
    private Workbook workbook;

    /**
     * 公式计算器
     */
    private FormulaEvaluator formulaEvaluator;

    /**
     * 当前读取到的行和列号
     * -1表示未读取到行号和列号
     * 这两种属性用于在读取数据时，获取当前读取到的行和列号，以提示错误信息和定位读取位置
     */
    @Getter
    private int currentReadRowIndex = -1;
    @Getter
    private int currentReadColumnIndex = -1;

    /**
     * 是否是事件驱动的读取
     */
    private boolean _eventDriven = false;

    public WorkBookContext(File file, DetectResult detectResult) {
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

    public boolean getEventDriven() {
        return _eventDriven;
    }

    public void setEventDriven() {
        this._eventDriven = true;
    }

    public void setCurrentReadRowIndex(int currentReadRowIndex) {
        this.currentReadRowIndex = currentReadRowIndex;
    }

    public void setCurrentReadColumnIndex(int currentReadColumnIndex) {
        this.currentReadColumnIndex = currentReadColumnIndex;
    }

    /**
     * 获取当前读取到的行和列号的可读字符串
     * @return 当前读取到的行和列号的可读字符串
     */
    public String getHumanReadablePosition() {
        return ExcelToolkit.getHumanReadablePosition(currentReadRowIndex, currentReadColumnIndex);
    }
}

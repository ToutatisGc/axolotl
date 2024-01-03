package cn.toutatis.xvoid.axolotl.excel;

import cn.toutatis.xvoid.axolotl.excel.support.AbstractContext;
import cn.toutatis.xvoid.axolotl.excel.support.tika.DetectResult;
import lombok.Getter;
import lombok.Setter;
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
     * 是否使用默认的读取配置
     */
    @Setter @Getter
    private boolean useDefaultReaderConfig = false;

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
     * [内部属性]
     * 直接指定读取的类
     * 在读取数据时使用不指定读取类型的读取方法时，使用该类读取数据
     */
    private Class<?> _directReadClass;
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

    protected Class<?> getDirectReadClass() {
        return _directReadClass;
    }

    protected void setDirectReadClass(Class<?> _directClass) {
        this._directReadClass = _directClass;
    }

    public boolean getEventDriven() {
        return _eventDriven;
    }

    protected void setEventDriven() {
        this._eventDriven = true;
    }

    protected void setCurrentReadRowIndex(int currentReadRowIndex) {
        this.currentReadRowIndex = currentReadRowIndex;
    }

    protected void setCurrentReadColumnIndex(int currentReadColumnIndex) {
        this.currentReadColumnIndex = currentReadColumnIndex;
    }

    /**
     * 获取当前读取到的行和列号的可读字符串
     * @return 当前读取到的行和列号的可读字符串
     */
    public String getCurrentHumanReadablePosition() {
        char i = (char) ( 'A' + currentReadColumnIndex);
        return String.format("%s", i+("%d".formatted(currentReadRowIndex + 1)));
    }
}

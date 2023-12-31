package cn.toutatis.xvoid.axolotl.excel;

import cn.toutatis.xvoid.axolotl.excel.support.AbstractContext;
import cn.toutatis.xvoid.axolotl.excel.support.tika.DetectResult;
import lombok.Getter;
import lombok.Setter;
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
     * 表内容缓存
     */
    private final Map<Integer, List<Row>> sheetData = new HashMap<>();

    /**
     * 当前读取到的行和列号
     * -1表示未读取到行号和列号
     * 这两种属性用于在读取数据时，获取当前读取到的行和列号，以提示错误信息和定位读取位置
     */
    private int currentReadRowIndex = -1;
    private int currentReadColumnIndex = -1;

    /**
     * [内部属性]
     * 直接指定读取的类
     * 在读取数据时使用不指定读取类型的读取方法时，使用该类读取数据
     */
    private Class<?> _directReadClass;
    //TODO 计划开发事件驱动
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

    protected Class<?> getDirectReadClass() {
        return _directReadClass;
    }

    protected void setDirectReadClass(Class<?> _directClass) {
        this._directReadClass = _directClass;
    }

    public boolean getEventDriven() {
        return _eventDriven;
    }

    public void setEventDriven(boolean _eventDriven) {
        this._eventDriven = _eventDriven;
    }
}

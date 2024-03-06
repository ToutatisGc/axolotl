package cn.toutatis.xvoid.axolotl.excel.reader;

import cn.toutatis.xvoid.axolotl.common.AbstractContext;
import cn.toutatis.xvoid.axolotl.excel.reader.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.tika.DetectResult;
import com.google.common.collect.HashBasedTable;
import com.google.common.io.ByteStreams;
import com.google.common.io.Files;
import lombok.Getter;
import lombok.Setter;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.InputStream;
import java.util.HashMap;
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
     * 工作簿的缓存数据
     */
    @Setter @Getter
    private byte[] dataCache;

    /**
     * 工作簿的表头缓存
     */
    @Getter
    private final Map<Integer, HashBasedTable<String,Integer,Integer>> headerCaches = new HashMap<>();

    /**
     * 转换器缓存
     */
    @Getter
    private Map<Class<?>, DataCastAdapter<?>> castAdapterCache = new HashMap<>();

    /**
     * 是否是事件驱动的读取
     */
    private boolean _eventDriven = false;

    @SneakyThrows
    public WorkBookContext(File file, DetectResult detectResult) {
        this.setFile(file);
        this.setDataCache(Files.toByteArray(file));
        this.setMimeType(detectResult.getCatchMimeType());
    }

    @SneakyThrows
    public WorkBookContext(InputStream ins, DetectResult detectResult) {
        this.setDataCache(ByteStreams.toByteArray(ins));
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

    /**
     * 获取指定索引的sheet
     * @param idx 索引
     * @return sheet
     */
    public Sheet getIndexSheet(int idx){
        return workbook.getSheetAt(idx);
    }
}

package cn.xvoid.axolotl.excel.writer.support.base;

import cn.xvoid.axolotl.excel.writer.components.widgets.Header;
import com.google.common.collect.HashBasedTable;
import lombok.Data;
import lombok.EqualsAndHashCode;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
@EqualsAndHashCode(callSuper = true)
public class AutoWriteContext extends WriteContext{

    /**
     * 工作薄实例
     */
    private Workbook workbook;

    /**
     * 表头信息
     */
    private Map<Integer,List<Header>> headers = new HashMap<>();

    /**
     * 表头行数(包含标题)
     */
    private Map<Integer,Integer> headerRowCount = new HashMap<>();

    /**
     * 数据
     */
    private List<?> datas;

    /**
     * 已经写入的行数
     * 注意:起始为-1
     */
    private Map<Integer,Integer> alreadyWriteRow = new HashMap<>();

    /**
     * 已经写入的列数
     * 注意:起始为0
     */
    private Map<Integer,Integer> alreadyWrittenColumns = new HashMap<>();

    /**
     * 当前写入数据序号
     * 注意:起始为-1
     */
    private Map<Integer,Integer> serialNumber = new HashMap<>();

    /**
     * 写入类信息
     */
    private Map<Integer,Class<?>> metaClass = new HashMap<>();

    /**
     * 表头列索引映射
     */
    private HashBasedTable<Integer,String,Integer> headerColumnIndexMapping = HashBasedTable.create();

    /**
     * 结尾合计
     */
    private HashBasedTable<Integer,Integer, BigDecimal> endingTotalMapping = HashBasedTable.create();

    /**
     * 执行结果
     */
    private List<AxolotlWriteResult> executeResults;

    public int getAndIncrementSerialNumber(){
        Integer serialNumber = alreadyWriteRow.getOrDefault(getSwitchSheetIndex(), -1);
        this.serialNumber.put(getSwitchSheetIndex(),serialNumber+1);
        return serialNumber;
    }

}

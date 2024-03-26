package cn.toutatis.xvoid.axolotl.excel.writer.support;

import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import com.google.common.collect.HashBasedTable;
import lombok.Data;
import lombok.EqualsAndHashCode;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.math.BigDecimal;
import java.util.List;

@Data
@EqualsAndHashCode(callSuper = true)
public class AutoWriteContext extends WriteContext{

    /**
     * 工作薄实例
     */
    private SXSSFWorkbook workbook;

    /**
     * 表头信息
     */
    private List<Header> headers;

    /**
     * 数据
     */
    private List<?> datas;

    /**
     * 已经写入的行数
     */
    private int alreadyWriteRow = -1;

    /**
     * 已经写入的列数
     */
    private int alreadyWrittenColumns = 0;

    /**
     * 当前写入数据序号
     */
    private int serialNumber = 1;

    /**
     * 写入类信息
     */
    private Class<?> metaClass;

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
        return serialNumber++;
    }

}

package cn.toutatis.xvoid.axolotl.excel.writer.support;

import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import lombok.Data;
import lombok.EqualsAndHashCode;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

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

}

package cn.toutatis.xvoid.axolotl.excel.writer.support;

import lombok.Data;
import lombok.EqualsAndHashCode;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

@Data
@EqualsAndHashCode(callSuper = true)
public class AutoWriteContext extends WriteContext{

    private SXSSFWorkbook workbook;

}

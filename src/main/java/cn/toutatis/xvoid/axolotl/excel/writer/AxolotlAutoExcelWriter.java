package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractInnerStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;

import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * 文档文件写入器
 * @author Toutatis_Gc
 */
public class AxolotlAutoExcelWriter extends AxolotlAbstractExcelWriter {

    /**
     * 日志工具
     * 日志记录器
     */
    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlAutoExcelWriter.class);

    private final AutoWriteConfig writeConfig;

    /**
     * 主构造函数
     *
     * @param autoWriteConfig 写入配置
     */
    public AxolotlAutoExcelWriter(AutoWriteConfig autoWriteConfig) {
        this.writeConfig = autoWriteConfig;
        this.workbook = this.initWorkbook(null);
        super.LOGGER = LOGGER;
    }

    @Override
    public AxolotlWriteResult write(Map<String, ?> header, List<?> circleDataList) throws AxolotlWriteException {
        LoggerHelper.info(LOGGER, writeContext.getCurrentWrittenBatchAndIncrement(writeConfig.getSheetIndex()));
        SXSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(writeConfig.getSheetIndex(), writeConfig.getSheetName());
        ExcelStyleRender styleRender = writeConfig.getStyleRender();
        if (styleRender instanceof AbstractInnerStyleRender innerStyleRender){
            innerStyleRender.setWriteConfig(writeConfig);
            innerStyleRender.renderHeader(sheet);
        }else {
            styleRender.renderHeader(sheet);
        }
        styleRender.renderData(sheet,circleDataList);
        return null;
    }

    @Override
    public void flush() {

    }

    @Override
    public void close() throws IOException {
        workbook.write(writeConfig.getOutputStream());
    }

    /**
     * 获取配置绑定索引
     */
    protected XSSFSheet getConfigBoundSheet() {
        return this.getWorkbookSheet(this.writeConfig.getSheetIndex());
    }
}

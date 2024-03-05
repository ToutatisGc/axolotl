package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AutoWriteContext;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;

import java.io.IOException;
import java.util.List;

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

    private final AutoWriteContext writeContext;

    /**
     * 主构造函数
     *
     * @param autoWriteConfig 写入配置
     */
    public AxolotlAutoExcelWriter(AutoWriteConfig autoWriteConfig) {
        this.writeConfig = autoWriteConfig;
        AutoWriteContext autoWriteContext = new AutoWriteContext();
        this.writeContext = autoWriteContext;
        super.writeContext = autoWriteContext;
        this.workbook = this.initWorkbook(null);
        super.LOGGER = LOGGER;
    }

    /**
     * 写入Excel数据
     * @param dataList 循环引用数据
     * @return 写入结果
     * @throws AxolotlWriteException 写入异常
     */
    public AxolotlWriteResult write(List<Header> headers, List<?> dataList) throws AxolotlWriteException {
        LoggerHelper.info(LOGGER, writeContext.getCurrentWrittenBatchAndIncrement(writeConfig.getSheetIndex()));
        SXSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(writeConfig.getSheetIndex(), writeConfig.getSheetName());
        ExcelStyleRender styleRender = writeConfig.getStyleRender();
        if (styleRender instanceof AbstractStyleRender innerStyleRender){
            innerStyleRender.setWriteConfig(writeConfig);
            innerStyleRender.renderHeader(sheet);
        }else {
            styleRender.renderHeader(sheet);
        }

        styleRender.renderData(sheet, dataList);
        return null;
    }

    public AxolotlWriteResult write(List<?> dataList) throws AxolotlWriteException {
        LoggerHelper.info(LOGGER, writeContext.getCurrentWrittenBatchAndIncrement(writeConfig.getSheetIndex()));
        SXSSFSheet sheet = workbook.createSheet();
        // TODO 渲染头部信息
        // TODO 解析实体注解
        // TODO 渲染实体数据到表
        // TODO 渲染结束数据
        workbook.setSheetName(writeConfig.getSheetIndex(), writeConfig.getSheetName());
        ExcelStyleRender styleRender = writeConfig.getStyleRender();
//        if (styleRender instanceof AbstractInnerStyleRender innerStyleRender){
//            innerStyleRender.setWriteConfig(writeConfig);
//            innerStyleRender.renderHeader(sheet);
//        }else {
//            styleRender.renderHeader(sheet);
//        }
        styleRender.renderData(sheet, dataList);
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

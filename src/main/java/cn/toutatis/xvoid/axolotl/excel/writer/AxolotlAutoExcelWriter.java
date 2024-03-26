package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AutoWriteContext;
import cn.toutatis.xvoid.axolotl.excel.writer.support.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

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
        super.LOGGER = LOGGER;
        this.writeConfig = autoWriteConfig;
        this.checkWriteConfig(this.writeConfig);
        AutoWriteContext autoWriteContext = new AutoWriteContext();
        this.workbook = this.initWorkbook(null);
        autoWriteContext.setWorkbook(this.workbook);
        this.writeContext = autoWriteContext;
        super.writeContext = autoWriteContext;
        writeContext.setSwitchSheetIndex(autoWriteConfig.getSheetIndex());
    }

    /**
     * 写入Excel数据
     * @param datas 循环引用数据
     * @return 写入结果
     * @throws AxolotlWriteException 写入异常
     */
    public AxolotlWriteResult write(List<Header> headers, List<?> datas) throws AxolotlWriteException {
        int switchSheetIndex = writeContext.getSwitchSheetIndex();
        info(LOGGER, writeContext.getCurrentWrittenBatchAndIncrement(switchSheetIndex));
        SXSSFSheet sheet;
        ExcelStyleRender styleRender = writeConfig.getStyleRender();
        if (styleRender == null){
            throw new AxolotlWriteException("请设置写入渲染器");
        }
        writeContext.getHeaders().put(switchSheetIndex,headers);
        writeContext.setDatas(datas);
        if (styleRender instanceof AbstractStyleRender innerStyleRender){
            innerStyleRender.setWriteConfig(writeConfig);
            innerStyleRender.setContext(writeContext);
        }
        if(writeContext.isFirstBatch(switchSheetIndex)){
            sheet = workbook.createSheet();
            writeContext.setWorkbook(workbook);
            styleRender.init(sheet);
            styleRender.renderHeader(sheet);
        }else {
            sheet = workbook.getSheetAt(switchSheetIndex);
        }
        styleRender.renderData(sheet, datas);
        return null;
    }

    public AxolotlWriteResult write(List<?> dataList) throws AxolotlWriteException {
        info(LOGGER, writeContext.getCurrentWrittenBatchAndIncrement(writeConfig.getSheetIndex()));
        SXSSFSheet sheet = workbook.createSheet();

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
        ExcelStyleRender styleRender = writeConfig.getStyleRender();
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            styleRender.finish(getWorkbook().getSheetAt(i));
        }
    }

    @Override
    public void close() throws IOException {
        OutputStream outputStream = writeConfig.getOutputStream();
        this.flush();
        workbook.write(outputStream);
        workbook.close();
    }

    /**
     * 获取配置绑定索引
     */
    protected XSSFSheet getConfigBoundSheet() {
        return this.getWorkbookSheet(this.writeConfig.getSheetIndex());
    }
}

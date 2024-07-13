package cn.xvoid.axolotl.excel.writer;

import cn.xvoid.axolotl.excel.writer.components.widgets.Header;
import cn.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.xvoid.axolotl.excel.writer.support.base.AutoWriteContext;
import cn.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.toolkit.validator.Validator;
import cn.xvoid.axolotl.toolkit.LoggerHelper;
import com.google.common.collect.Lists;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.slf4j.Logger;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import static cn.xvoid.axolotl.toolkit.LoggerHelper.info;

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

    /**
     * 写入配置
     */
    private final AutoWriteConfig writeConfig;

    /**
     * 写入上下文
     */
    private final AutoWriteContext writeContext;

    /**
     * 主构造函数
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
        if (styleRender instanceof AbstractStyleRender){
            AbstractStyleRender innerStyleRender = (AbstractStyleRender)styleRender;
            innerStyleRender.setWriteConfig(writeConfig);
            innerStyleRender.setContext(writeContext);
            innerStyleRender.getComponentRender().setConfig(writeConfig);
            innerStyleRender.getComponentRender().setContext(writeContext);
        }
        if(writeContext.isFirstBatch(switchSheetIndex)){
            sheet = workbook.createSheet();
            writeContext.setWorkbook(workbook);
            styleRender.init(sheet);
            styleRender.renderHeader(sheet);
        }else {
            sheet = workbook.getSheetAt(switchSheetIndex);
        }
        if(datas != null){
            if (Validator.objNotNull(datas)){
                writeConfig.setMetaClass(datas.get(0).getClass());
                writeConfig.autoProcessEntity2OpenDictPolicy();
            }
            styleRender.renderData(sheet, datas);
        }
        return null;
    }

    /**
     * 仅写入列表数据
     * @param data 列表数据
     * @return 写入结果
     * @throws AxolotlWriteException 写入异常
     */
    public AxolotlWriteResult write(List<?> data) throws AxolotlWriteException {
        return this.write(Lists.newArrayList(),data);
    }

    public AxolotlWriteResult write(Class<?> metaClass,List<?> data) throws AxolotlWriteException {
        return this.write(ExcelToolkit.getHeaderList(metaClass),data);
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

}

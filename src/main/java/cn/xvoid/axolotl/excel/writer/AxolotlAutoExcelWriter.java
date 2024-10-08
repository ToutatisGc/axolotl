package cn.xvoid.axolotl.excel.writer;

import cn.xvoid.axolotl.excel.writer.components.widgets.AxolotlImage;
import cn.xvoid.axolotl.excel.writer.components.widgets.Header;
import cn.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.xvoid.axolotl.excel.writer.support.base.AutoWriteContext;
import cn.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
import cn.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.toolkit.validator.Validator;
import cn.xvoid.axolotl.toolkit.LoggerHelper;
import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.Sheet;
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
        LoggerHelper.info(LOGGER, writeContext.getCurrentWrittenBatchAndIncrement(switchSheetIndex));
        Sheet sheet;
        ExcelStyleRender styleRender = writeConfig.getStyleRender();
        if (styleRender == null){
            throw new AxolotlWriteException("请设置写入渲染器");
        }
        writeContext.getHeaders().put(switchSheetIndex,headers);
        writeContext.setDatas(datas);
        if (styleRender instanceof AbstractStyleRender innerStyleRender){
            innerStyleRender.setWriteConfig(writeConfig);
            innerStyleRender.setContext(writeContext);
            innerStyleRender.getComponentRender().setConfig(writeConfig);
            innerStyleRender.getComponentRender().setContext(writeContext);
        }
        if(writeContext.isFirstBatch(switchSheetIndex)){
            sheet = ExcelToolkit.createOrCatchSheet(workbook, switchSheetIndex);
            writeContext.setWorkbook(workbook);
            styleRender.init(sheet);
            styleRender.renderHeader(sheet);
        }else {
            sheet = ExcelToolkit.createOrCatchSheet(workbook, switchSheetIndex);
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
        ExcelStyleRender styleRender = getWriteConfig().getStyleRender();
        int numberOfSheets = workbook.getNumberOfSheets();
        for (int i = 0; i < numberOfSheets; i++) {
            styleRender.finish(getWorkbook().getSheetAt(i));
        }
    }

    @Override
    public AutoWriteConfig getWriteConfig() {
        return writeConfig;
    }

    @Override
    public void close() throws IOException {
        OutputStream outputStream = writeConfig.getOutputStream();
        this.flush();
        workbook.write(outputStream);
        workbook.close();
    }

    @Override
    public void writeImage(int sheetIndex, AxolotlImage axolotlImage) {
        ExcelStyleRender styleRender = writeConfig.getStyleRender();
        Sheet sheet = ExcelToolkit.createOrCatchSheet(getWorkbook(), sheetIndex);
        if (styleRender == null){
            throw new AxolotlWriteException("请设置写入渲染器");
        }
        if (styleRender instanceof AbstractStyleRender innerStyleRender){
            innerStyleRender.setWriteConfig(writeConfig);
            innerStyleRender.setContext((AutoWriteContext) writeContext);
            innerStyleRender.getComponentRender().setConfig(writeConfig);
            innerStyleRender.getComponentRender().setContext(writeContext);
            if(writeContext.isFirstBatch(sheetIndex)){
                ((AutoWriteContext)writeContext).setWorkbook(workbook);
                innerStyleRender.init(sheet);
                innerStyleRender.renderHeader(sheet);
            }
        }
        super.writeImage(sheetIndex, axolotlImage);
    }
}

package cn.toutatis.xvoid.axolotl.excel.reader;

import cn.toutatis.xvoid.axolotl.Meta;
import cn.toutatis.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.reader.support.AxolotlAbstractExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException.ExceptionType;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkitKt;
import org.slf4j.Logger;

import java.io.File;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class AxolotlExcelReader<T> extends AxolotlAbstractExcelReader<T> implements Iterator<List<T>> {

    /**
     * 日志工具
     */
    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlExcelReader.class);

    /**
     * 当前读取批次
     */
    private int currentReadBatch = -1;

    public AxolotlExcelReader(File excelFile) {
        super(excelFile);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(File excelFile, boolean withDefaultConfig) {
        super(excelFile, withDefaultConfig);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(File excelFile, Class<T> clazz) {
        super(excelFile, clazz);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(InputStream ins) {
        super(ins);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(InputStream ins, Class<T> clazz) {
        super(ins, clazz);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(File excelFile, Class<T> clazz, boolean withDefaultConfig) {
        super(excelFile, clazz, withDefaultConfig);
        super.LOGGER = LOGGER;
    }

    /**
     * 是否有下一批数据
     */
    @Override
    public boolean hasNext() {
        return currentReadBatch * AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE < getRecordRowNumber();
    }

    /**
     * 获取下一批数据
     */
    @Override
    public List<T> next() {
        if (!hasNext()){
            throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR,"读取数据错误");
        }
        currentReadBatch++;
        LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME,"读取数据行数:"+currentReadBatch*AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE);
        return this.readSheetData(
                currentReadBatch*AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE,
                (currentReadBatch+1)*AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE
        );
    }
}

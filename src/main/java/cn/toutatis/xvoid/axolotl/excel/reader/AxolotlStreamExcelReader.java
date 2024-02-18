package cn.toutatis.xvoid.axolotl.excel.reader;

import cn.toutatis.xvoid.axolotl.excel.reader.support.AxolotlAbstractExcelReader;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import org.slf4j.Logger;

import java.io.File;
import java.io.InputStream;

public class AxolotlStreamExcelReader<T> extends AxolotlAbstractExcelReader<T> {

    /**
     * 日志工具
     */
    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlExcelReader.class);

    public AxolotlStreamExcelReader(File excelFile) {
        super(excelFile);
        super.LOGGER = LOGGER;
    }

    public AxolotlStreamExcelReader(File excelFile, boolean withDefaultConfig) {
        super(excelFile, withDefaultConfig);
        super.LOGGER = LOGGER;
    }

    public AxolotlStreamExcelReader(File excelFile, Class<T> clazz) {
        super(excelFile, clazz);
        super.LOGGER = LOGGER;
    }

    public AxolotlStreamExcelReader(InputStream ins) {
        super(ins);
        super.LOGGER = LOGGER;
    }

    public AxolotlStreamExcelReader(InputStream ins, Class<T> clazz) {
        super(ins, clazz);
        super.LOGGER = LOGGER;
    }

    public AxolotlStreamExcelReader(File excelFile, Class<T> clazz, boolean withDefaultConfig) {
        super(excelFile, clazz, withDefaultConfig);
        super.LOGGER = LOGGER;
    }

    @Override
    protected void loadFileDataToWorkBook() {
        super.loadFileDataToWorkBook();
    }
}

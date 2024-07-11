package cn.toutatis.xvoid.axolotl.excel.reader;

import cn.toutatis.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.toutatis.xvoid.axolotl.excel.reader.support.AxolotlAbstractExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.support.AxolotlReadInfo;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.axolotl.excel.reader.support.stream.AxolotlExcelStream;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.toolkit.tika.TikaShell;
import cn.xvoid.toolkit.clazz.ReflectToolkit;
import cn.xvoid.toolkit.log.LoggerToolkit;
import com.github.pjfanning.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.RecordFormatException;
import org.slf4j.Logger;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;

/**
 * Excel流形式读取器
 * <p>大文件读取时，推荐使用流形式读取器，相比于AxolotlExcelReader内存占用实际上没什么区别，但是读取速度更快。</p>
 * <p>对于AxolotlStreamExcelReader读取数据时，少了很多特性支持，不支持分页等，因为流形式读取器读取时，不会将整个文件读入内存。</p>
 * <p>流支持基于excel-streaming-reader,并且该流读取器只支持<b>xlsx</b>格式</p>
 * <p>excel-streaming-reader在底层使用了一些Apache POI代码。该代码在处理xlsx时使用内存和/或临时文件来存储临时数据。对于非常大的文件，您可能希望使用临时文件。</p>
 * <p>对于StreamingReader.builder()，不要设置setAvoidTempFiles(true)。您还应该考虑调优POI设置。具体来说，请考虑设置这些属性:</p>
 * <p>ZipInputStreamZipEntrySource.setThresholdBytesForTempFiles(16384); //16KB</p>
 * <p>ZipPackage.setUseTempFilePackageParts(true);</p>
 * @param <T> 转换实体
 * @author Toutatis_Gc
 * @since 0.0.9-ALPHA-1
 */
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

    /**
     * 检查文件类型
     * 流方式仅支持xlsx格式
     * @param file 工作簿文件
     * @param ins 输入流
     * @return 文件检测结果
     */
    @Override
    protected DetectResult checkFileFormat(File file, InputStream ins) {
        return this.getFileOrStreamDetectResult(file, ins, TikaShell.OOXML_EXCEL);
    }

    /**
     * 加载大文件
     */
    @Override
    protected void loadFileDataToWorkBook() {
        try(InputStream fis = new ByteArrayInputStream(workBookContext.getDataCache())){
            Workbook workbook = StreamingReader
                    .builder()
                    .rowCacheSize(1000)
                    .bufferSize(4096)
                    .open(fis);
            workBookContext.setWorkbook(workbook);
        } catch (IOException | RecordFormatException e) {
            LOGGER.error("加载文件失败",e);
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,e.getMessage());
        }
    }

    /**
     * 读取行数据转换为对象
     * @param row 行
     * @param readerConfig 读取配置
     * @return 对象
     * @param <RT> 对象类型
     */
    public <RT> RT castRow2Instance(Row row, ReaderConfig<RT> readerConfig){
        RT instance = readerConfig.getCastClassInstance();
        if (ExcelToolkit.blankRowCheck(row,readerConfig)){
            if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.INCLUDE_EMPTY_ROW)){
                return instance;
            }else{
                return null;
            }
        }
        this.convertCellToInstance(row,instance,readerConfig);
        String needRecordInfo = readerConfig.getNeedRecordInfo();
        if (instance != null && needRecordInfo != null ){
            Sheet sheet = workBookContext.getIndexSheet(readerConfig.getSheetIndex());
            try {
                Field field = readerConfig.getCastClass().getDeclaredField(needRecordInfo);
                AxolotlReadInfo axolotlReadInfo = new AxolotlReadInfo();
                axolotlReadInfo.setSheetIndex(readerConfig.getSheetIndex());
                axolotlReadInfo.setSheetName(sheet.getSheetName());
                axolotlReadInfo.setRowNumber(row.getRowNum());
                ReflectToolkit.setObjectField(instance,field,axolotlReadInfo);
            } catch (NoSuchFieldException e) {
                throw new RuntimeException(e);
            }
        }
        return instance;
    }

    /**
     * 数据迭代器
     * @param readerConfig 读取配置
     * @param <RT> 对象类型
     * @return 数据迭代器
     */
    public <RT> AxolotlExcelStream<RT> dataIterator(ReaderConfig<RT> readerConfig){
        this.searchSheet(readerConfig);
        this.preCheckAndFixReadConfig(readerConfig);
        return new AxolotlExcelStream<>(this, readerConfig);
    }

}

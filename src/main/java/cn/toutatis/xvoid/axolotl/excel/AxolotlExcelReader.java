package cn.toutatis.xvoid.axolotl.excel;

import cn.toutatis.xvoid.axolotl.excel.support.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.excel.support.tika.TikaShell;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.util.RecordFormatException;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class AxolotlExcelReader<T extends Object> {

    /**
     * 日志
     */
    private final Logger LOGGER  = LoggerToolkit.getLogger(AxolotlExcelReader.class);

    /**
     * 工作簿元信息
     */
    @Getter
    private WorkBookContext workBookContext;

    /**
     * 构造文件读取器
     */
    public AxolotlExcelReader(File excelFile) {
        this(excelFile,true);
    }

    @SuppressWarnings("unchecked")
    public AxolotlExcelReader(File excelFile, boolean withDefaultConfig) {
        this(excelFile, (Class<T>) Object.class,withDefaultConfig);
    }

    public AxolotlExcelReader(File excelFile, Class<T> clazz) {
        this(excelFile,clazz,true);
    }

    /**
     * [ROOT]
     * 构造文件读取器
     * @param excelFile Excel工作簿文件
     * @param withDefaultConfig 是否使用默认配置
     */
    public AxolotlExcelReader(File excelFile, Class<T> clazz, boolean withDefaultConfig) {
        if (clazz == null){
            throw new IllegalArgumentException("读取的类型对象不能为空");
        }
        this.detectFileAndInitWorkbook(excelFile);
        this.workBookContext.setDirectReadClass(clazz);
        this.workBookContext.setUseDefaultReaderConfig(withDefaultConfig);
    }

    /**
     * 初始化读取Excel文件
     * 1.初始化加载文件先判断文件是否正常并且是需要的格式
     * 2.将文件加载到POI工作簿中
     * @param excelFile Excel工作簿文件
     */
    private void detectFileAndInitWorkbook(File excelFile) {
        // 检查文件格式是否为XLSX
        DetectResult detectResult = TikaShell.detect(excelFile, TikaShell.OOXML_EXCEL,false);
        if (!detectResult.isDetect()){
            DetectResult.FileStatus currentFileStatus = detectResult.getCurrentFileStatus();
            if (currentFileStatus == DetectResult.FileStatus.FILE_MIME_TYPE_PROBLEM ||
                    currentFileStatus == DetectResult.FileStatus.FILE_SUFFIX_PROBLEM
            ){
                //如果是因为后缀不匹配或媒体类型不匹配导致识别不通过 换XLS格式再次识别
                detectResult = TikaShell.detect(excelFile, TikaShell.MS_EXCEL,false);
            }else {
                //如果是预检查不通过  抛出异常
                detectResult.throwException();
            }
        }
        // 检查文件是否正常并且是需要的类型，否则抛出异常
        if (detectResult.isDetect()){
            workBookContext = new WorkBookContext(excelFile,detectResult);
        }else{
            detectResult.throwException();
        }
        // 读取文件加载到元信息
        try(FileInputStream fis = new FileInputStream(workBookContext.getFile())){
            Workbook workbook = WorkbookFactory.create(fis);
            workBookContext.setWorkbook(workbook);
        } catch (IOException | RecordFormatException e) {
            LOGGER.error("加载文件失败",e);
            throw new RuntimeException(e);
        }
    }

}

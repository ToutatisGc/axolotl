package cn.toutatis.xvoid.axolotl.excel;

import cn.toutatis.xvoid.axolotl.excel.support.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.excel.support.tika.TikaShell;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class GracefulExcelReader<T> {

    /**
     * 日志
     */
    private final Logger LOGGER  = LoggerToolkit.getLogger(GracefulExcelReader.class);

    /**
     * 工作簿元信息
     */
    @Getter
    private WorkBookContext workBookContext;

    /**
     * 构造文件读取器
     */
    public GracefulExcelReader(File excelFile) {
        this(excelFile,true);
    }


    /**
     * 构造文件读取器
     *
     * @param excelFile Excel工作簿文件
     * @param withDefaultConfig 是否使用默认配置
     */
    public GracefulExcelReader(File excelFile,boolean withDefaultConfig) {
        this.initWorkbook(excelFile);
        this.workBookContext.setUseDefaultReaderConfig(withDefaultConfig);
    }

    public GracefulExcelReader(File excelFile, Class<T> clazz) {
        this(excelFile,clazz,true);
    }

    public GracefulExcelReader(File excelFile, Class<T> clazz,boolean withDefaultConfig) {
        this.initWorkbook(excelFile);
        if (clazz == null){
            throw new IllegalArgumentException("读取的类型对象不能为空");
        }
        this.workBookContext.setDirectReadClass(clazz);
        this.workBookContext.setUseDefaultReaderConfig(withDefaultConfig);
    }

    /**
     * 初始化读取Excel文件
     * 1.初始化加载文件先判断文件是否正常并且是需要的格式
     * 2.将文件加载到POI工作簿中
     * @param excelFile Excel工作簿文件
     */
    private void initWorkbook(File excelFile) {
        // 检查文件是否正常
        TikaShell.preCheckFileNormalThrowException(excelFile);
        DetectResult detectResult = TikaShell.detect(excelFile, TikaShell.OOXML_EXCEL,true);
        if (!detectResult.isDetect()){
            // 没有识别到XLSX格式再尝试识别XLS格式
            DetectResult.FileStatus currentFileStatus = detectResult.getCurrentFileStatus();
            if (currentFileStatus == DetectResult.FileStatus.FILE_MIME_TYPE_PROBLEM ||
                    currentFileStatus == DetectResult.FileStatus.FILE_SUFFIX_PROBLEM
            ){
                detectResult = TikaShell.detect(excelFile, TikaShell.MS_EXCEL,true);
            }else {
                detectResult.throwException();
            }
        }
        // 检查文件是否正常并且是需要的类型，否则抛出异常
        if (detectResult.isDetect() && detectResult.isWantedMimeType()){
            workBookContext = new WorkBookContext(excelFile,detectResult);
        }else{
            detectResult.throwException();
        }
        // 读取文件加载到元信息
        try(FileInputStream fis = new FileInputStream(workBookContext.getFile())){
            Workbook workbook = WorkbookFactory.create(fis);
            workBookContext.setWorkbook(workbook);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}

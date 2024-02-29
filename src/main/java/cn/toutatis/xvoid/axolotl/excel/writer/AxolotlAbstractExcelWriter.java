package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.common.CommonMimeType;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.support.WriteContext;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.axolotl.toolkit.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.toolkit.tika.TikaShell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.debug;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.format;

public abstract class AxolotlAbstractExcelWriter implements AxolotlExcelWriter{

    /**
     * 日志工具
     * 日志记录器
     */
    protected Logger LOGGER;

    /**
     * 由文件加载而来的工作簿文件信息
     * 写入工作簿
     */
    protected SXSSFWorkbook workbook;

    /**
     * 写入上下文
     */
    protected final WriteContext writeContext = new WriteContext();

    /**
     * 写入配置
     */
    protected final WriterConfig writerConfig;

    public AxolotlAbstractExcelWriter(WriterConfig writerConfig) {
        this.writerConfig = writerConfig;
    }

    /**
     * 初始化工作簿
     *
     * @param templateFile 模板文件
     * @return 工作簿
     */
    protected SXSSFWorkbook initWorkbook(File templateFile) {
        SXSSFWorkbook workbook;
        // 读取模板文件内容
        if (templateFile != null){
            debug(LOGGER, format("正在使用模板文件[%s]作为写入模板",templateFile.getAbsolutePath()));
            TikaShell.preCheckFileNormalThrowException(templateFile);
            DetectResult detect = TikaShell.detect(templateFile, CommonMimeType.OOXML_EXCEL);
            if (!detect.isWantedMimeType()){
                throw new AxolotlWriteException("请使用xlsx文件作为写入模板");
            }
            this.writeContext.setFile(templateFile);
            try (FileInputStream fis = new FileInputStream(templateFile)){
                OPCPackage opcPackage = OPCPackage.open(fis);
                workbook = new SXSSFWorkbook(XSSFWorkbookFactory.createWorkbook(opcPackage));
            }catch (IOException | InvalidFormatException e){
                e.printStackTrace();
                throw new AxolotlWriteException(format("模板文件[%s]读取失败",templateFile.getAbsolutePath()));
            }
        }else {
            workbook = new SXSSFWorkbook();
        }
        return workbook;
    }

    /**
     * 获取配置绑定索引
     */
    protected XSSFSheet getConfigBoundSheet() {
        return this.getWorkbookSheet(this.writerConfig.getSheetIndex());
    }

    /**
     * 获取工作簿对应的工作表
     *
     * @param sheetIndex 工作表索引
     * @return 工作表
     */
    protected XSSFSheet getWorkbookSheet(int sheetIndex) {
        XSSFSheet sheet = workbook.getXSSFWorkbook().getSheetAt(sheetIndex);
        if (sheet == null){
            throw new AxolotlWriteException(LoggerHelper.format("工作簿索引[%s]对应的工作表不存在",this.writerConfig.getSheetIndex()));
        }
        return sheet;
    }


}

package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.common.CommonMimeType;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.WriteContext;
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
import java.io.OutputStream;

import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;

/**
 * 抽象工作簿写入器
 */
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
    protected WriteContext writeContext;

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
     * 获取工作簿对应的工作表
     *
     * @param sheetIndex 工作表索引
     * @return 工作表
     */
    protected XSSFSheet getWorkbookSheet(int sheetIndex) {
        XSSFSheet sheet = workbook.getXSSFWorkbook().getSheetAt(sheetIndex);
        if (sheet == null){
            throw new AxolotlWriteException(LoggerHelper.format("工作簿索引[%s]对应的工作表不存在",sheetIndex));
        }
        return sheet;
    }

    @Override
    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

    @Override
    public void switchSheet(int sheetIndex) {
        LoggerHelper.debug(LOGGER,"切换到工作表[%s]",sheetIndex);
//        ExcelToolkit.s
        // TODO 创建工作表
        this.writeContext.setSwitchSheetIndex(sheetIndex);
    }

    /**
     * 检查写入配置
     * @param writeConfig 写入配置
     */
    protected void checkWriteConfig(CommonWriteConfig writeConfig){
        if(writeConfig == null){
            String message = "写入配置不能为空";
            error(LOGGER,message);
            throw new AxolotlWriteException(message);
        }
        OutputStream outputStream = writeConfig.getOutputStream();
        if(outputStream == null){
            String message = "输出流不能为空,请指定输出流";
            error(LOGGER,message);
            throw new AxolotlWriteException(message);
        }
    }

}

package cn.xvoid.axolotl.excel.writer;

import cn.hutool.core.util.IdUtil;
import cn.xvoid.axolotl.common.CommonMimeType;
import cn.xvoid.axolotl.excel.writer.components.widgets.AxolotlImage;
import cn.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.xvoid.axolotl.excel.writer.style.AbstractStyleRender;
import cn.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.xvoid.axolotl.excel.writer.support.base.AutoWriteContext;
import cn.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
import cn.xvoid.axolotl.excel.writer.support.base.WriteContext;
import cn.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.xvoid.axolotl.toolkit.LoggerHelper;
import cn.xvoid.axolotl.toolkit.tika.DetectResult;
import cn.xvoid.axolotl.toolkit.tika.TikaShell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;

import java.io.*;
import java.util.Base64;

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
    protected Workbook workbook;

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
    protected Workbook initWorkbook(File templateFile) {
        Workbook workbook;
        // 读取模板文件内容
        if (templateFile != null){
            LoggerHelper.debug(LOGGER, LoggerHelper.format("正在使用模板文件[%s]作为写入模板",templateFile.getAbsolutePath()));
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
                throw new AxolotlWriteException(LoggerHelper.format("模板文件[%s]读取失败",templateFile.getAbsolutePath()));
            }
        }else {
            workbook = new SXSSFWorkbook();
        }
        return workbook;
    }

    public void writeImage(int sheetIndex,AxolotlImage axolotlImage){
        if (axolotlImage == null){throw new AxolotlWriteException("图片对象不能为空");}
        axolotlImage.checkImage();
        CommonWriteConfig writeConfig = getWriteConfig();
        Sheet workbookSheet;
        workbookSheet = ExcelToolkit.createOrCatchSheet(workbook, sheetIndex);
        if (AutoWriteConfig.class.equals(writeConfig.getClass())){
            AutoWriteConfig autoWriteConfig = (AutoWriteConfig) writeConfig;
            ExcelStyleRender styleRender = autoWriteConfig.getStyleRender();
            if (styleRender == null){
                throw new AxolotlWriteException("请设置写入渲染器");
            }
            if (styleRender instanceof AbstractStyleRender innerStyleRender){
                innerStyleRender.setWriteConfig(autoWriteConfig);
                innerStyleRender.setContext((AutoWriteContext) writeContext);
                innerStyleRender.getComponentRender().setConfig(writeConfig);
                innerStyleRender.getComponentRender().setContext(writeContext);
                if(writeContext.isFirstBatch(sheetIndex)){
                    ((AutoWriteContext)writeContext).setWorkbook(workbook);
                    innerStyleRender.init(workbookSheet);
                    innerStyleRender.renderHeader(workbookSheet);
                }
            }
        }
        int pictureIndex = workbook.addPicture(axolotlImage.getData(), axolotlImage.getImageFormat());
        CreationHelper helper = workbook.getCreationHelper();
        ClientAnchor clientAnchor = helper.createClientAnchor();
        clientAnchor.setAnchorType(axolotlImage.getAnchorType());
        clientAnchor.setCol1(axolotlImage.getStartColumn());
        clientAnchor.setRow1(axolotlImage.getStartRow());
        clientAnchor.setCol2(axolotlImage.getEndColumn());
        clientAnchor.setRow2(axolotlImage.getEndRow());
        workbookSheet.createDrawingPatriarch().createPicture(clientAnchor, pictureIndex);
    }


    /**
     * 获取工作簿对应的工作表
     *
     * @param sheetIndex 工作表索引
     * @return 工作表
     */
    protected Sheet getWorkbookSheet(int sheetIndex) {
        Sheet sheet;
        if (workbook.getClass() == SXSSFWorkbook.class){
            sheet = ((SXSSFWorkbook) workbook).getXSSFWorkbook().getSheetAt(sheetIndex);
        }else {
            sheet = ((XSSFWorkbook) workbook).getSheetAt(sheetIndex);
        }
        if (sheet == null){
            throw new AxolotlWriteException(LoggerHelper.format("工作簿索引[%s]对应的工作表不存在",sheetIndex));
        }
        return sheet;
    }

    protected int getSheetIndex(Sheet sheet){
        int sheetIndex;
        if (workbook.getClass() == SXSSFWorkbook.class){
            sheetIndex = ((SXSSFWorkbook) workbook).getXSSFWorkbook().getSheetIndex(sheet);
        }else {
            sheetIndex = ((XSSFWorkbook) workbook).getSheetIndex(sheet);
        }
        return sheetIndex;
    }

    @Override
    public Workbook getWorkbook() {
        return workbook;
    }

    @Override
    public void switchSheet(int sheetIndex) {
        LoggerHelper.debug(LOGGER,"切换到工作表[%s]",sheetIndex);
        this.writeContext.setSwitchSheetIndex(sheetIndex);
    }

    /**
     * 检查写入配置
     * @param writeConfig 写入配置
     */
    protected void checkWriteConfig(CommonWriteConfig writeConfig){
        if(writeConfig == null){
            String message = "写入配置不能为空";
            LoggerHelper.error(LOGGER,message);
            throw new AxolotlWriteException(message);
        }
        OutputStream outputStream = writeConfig.getOutputStream();
        if(outputStream == null){
            String message = "输出流不能为空,请指定输出流";
            LoggerHelper.error(LOGGER,message);
            throw new AxolotlWriteException(message);
        }
    }

}

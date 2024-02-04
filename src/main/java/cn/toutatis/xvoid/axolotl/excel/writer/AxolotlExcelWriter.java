package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.common.CommonMimeType;
import cn.toutatis.xvoid.axolotl.excel.writer.constant.TemplatePlaceholderPattern;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractInnerStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.CellAddress;
import cn.toutatis.xvoid.axolotl.excel.writer.support.WriteContext;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.axolotl.toolkit.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.toolkit.tika.TikaShell;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import com.alibaba.fastjson.JSONObject;
import com.google.common.collect.HashBasedTable;
import lombok.SneakyThrows;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 文档文件写入器
 * @author Toutatis_Gc
 */
public class AxolotlExcelWriter {

    /**
     * 日志记录器
     */
    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlExcelWriter.class);

    /**
     * 写入工作簿
     */
    private final SXSSFWorkbook workbook;

    /**
     * 写入上下文
     */
    private final WriteContext writeContext = new WriteContext();

    /**
     * 写入配置
     */
    private final WriterConfig writerConfig;

    /**
     * 主构造函数
     * @param writerConfig 写入配置
     */
    public AxolotlExcelWriter(WriterConfig writerConfig) {
        this.writerConfig = writerConfig;
        this.workbook = this.initWorkbook(null);
    }

    /**
     * 构造函数
     * 可以写入一个模板文件
     * @param templateFile 模板文件
     * @param writerConfig 写入配置
     */
    public AxolotlExcelWriter(File templateFile, WriterConfig writerConfig) {
        TikaShell.preCheckFileNormalThrowException(templateFile);
        this.writerConfig = writerConfig;
        this.workbook = this.initWorkbook(templateFile);
    }

    /**
     * 初始化工作簿
     * @param templateFile 模板文件
     * @return 工作簿
     */
    private SXSSFWorkbook initWorkbook(File templateFile) {
        SXSSFWorkbook workbook;
        // 读取模板文件内容
        if (templateFile != null){
            LoggerHelper.debug(LOGGER,LoggerHelper.format("正在使用模板文件[%s]作为写入模板",templateFile.getAbsolutePath()));
            TikaShell.preCheckFileNormalThrowException(templateFile);
            DetectResult detect = TikaShell.detect(templateFile, CommonMimeType.OOXML_EXCEL);
            if (!detect.isWantedMimeType()){
                throw new AxolotlWriteException("请使用xlsx文件作为写入模板");
            }
            this.writeContext.setTemplateFile(templateFile);
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

    /**
     * 关闭工作簿所对应输出流
     * @throws IOException IO异常
     */
    public void close() throws IOException {
        workbook.close();
        writerConfig.getOutputStream().close();
    }

    @SneakyThrows
    public void write(){
        SXSSFSheet sheet = workbook.createSheet();
        workbook.setSheetName(writerConfig.getSheetIndex(),writerConfig.getSheetName());
        ExcelStyleRender styleRender = writerConfig.getStyleRender();
        if (styleRender instanceof AbstractInnerStyleRender innerStyleRender){
            innerStyleRender.setWriterConfig(writerConfig);
            innerStyleRender.renderHeader(sheet);
        }else {
            styleRender.renderHeader(sheet);
        }
        styleRender.renderData(sheet,writerConfig.getData());
        workbook.write(writerConfig.getOutputStream());
    }

    /**
     * 写入数据到模板
     * @param sheetIndex 工作簿索引
     * @param singleMap 固定值映射
     * @param data 循环数据
     */
    @SneakyThrows
    public void writeToTemplate(int sheetIndex, Map<String,?> singleMap, List<Object> data){
        XSSFSheet sheet;
        if (writeContext.isTemplateWrite()){
            sheet = workbook.getXSSFWorkbook().getSheetAt(sheetIndex);
            if (sheet != null){
                this.resolveTemplate(sheet);
                HashBasedTable<Integer, String, CellAddress> singleReferenceData = this.writeContext.getSingleReferenceData();
                Map<String, CellAddress> singleAddressMapping = singleReferenceData.row(sheetIndex);
                for (String singleKey : singleAddressMapping.keySet()) {
                    CellAddress cellAddress = singleAddressMapping.get(singleKey);
                    XSSFCell cell = sheet.getRow(cellAddress.getRowPosition()).getCell(cellAddress.getColumnPosition());
                    if (singleMap!= null && singleMap.containsKey(singleKey)){
                        cell.setCellValue(cellAddress.replacePlaceholder(singleMap.get(singleKey).toString()));
                    }else {
                        cell.setCellValue(cellAddress.replacePlaceholder(""));
                    }
                }
                // TODO 写入循环数据
                Map<String, CellAddress> circleReferenceData = this.writeContext.getCircleReferenceData().row(sheetIndex);
                // TODO 列漂移写入
//                Collections.max()
                workbook.write(writerConfig.getOutputStream());
            }else{
                throw new AxolotlWriteException(LoggerHelper.format("工作表索引[%s]模板中不存在",sheetIndex));
            }
        }else{
            throw new AxolotlWriteException("非模板写入请使用write方法");
        }
    }

    /**
     * 解析数据实体类型
     * @param data 数据集合
     */
    private void resolveDataField(List<Object> data){
        if (data == null || data.isEmpty()){
            return;
        }
        Object dataTmp = data.get(0);
        Class<?> dataClass = dataTmp.getClass();

    }

    /**
     * 解析模板
     * @param sheet 工作表
     */
    private void resolveTemplate(Sheet sheet){
        int lastRowNum = sheet.getLastRowNum();
        for (int rowIdx = 0; rowIdx <= lastRowNum; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row != null){
                short lastCellNum = row.getLastCellNum();
                for (int cellIdx = 0; cellIdx < lastCellNum; cellIdx++) {
                    Cell cell = row.getCell(cellIdx);
                    if (cell != null && CellType.STRING.equals(cell.getCellType())){
                        String cellValue = cell.getStringCellValue();
                        int sheetIndex = workbook.getXSSFWorkbook().getSheetIndex(sheet);
                        CellAddress cellAddress = new CellAddress(cellValue,rowIdx, cellIdx,cell.getCellStyle());
                        findPlaceholderData(
                                writeContext.getSingleReferenceData(),
                                TemplatePlaceholderPattern.SINGLE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cellAddress
                        );
                        findPlaceholderData(
                                writeContext.getCircleReferenceData(),
                                TemplatePlaceholderPattern.CIRCLE_REFERENCE_TEMPLATE_PATTERN, sheetIndex, cellAddress
                        );
                    }
                }
            }
        }
        LoggerHelper.debug(LOGGER,LoggerHelper.format("解析模板完成，共解析到%s个占位符",writeContext.getSingleReferenceData().size() + writeContext.getCircleReferenceData().size()));
    }

    /**
     * 解析模板值到变量
     * @param referenceData 引用数据
     * @param pattern 模板匹配正则
     * @param sheetIndex 工作簿索引
     * @param cellAddress 单元格地址
     */
    private void findPlaceholderData(HashBasedTable<Integer, String, CellAddress> referenceData, Pattern pattern, int sheetIndex, CellAddress cellAddress) {
        Matcher matcher = pattern.matcher(cellAddress.getCellValue());
        if (matcher.find()) {
            cellAddress.setPlaceholder(matcher.group());
            referenceData.put(sheetIndex, matcher.group(1), cellAddress);
        }
    }

    public static void main(String[] args) throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\" + IdUtil.randomUUID() + ".xlsx"));

        WriterConfig writerConfig = new WriterConfig();
        writerConfig.setTitle("测试生成表标题");
        ArrayList<String> columnNames = new ArrayList<>();
        columnNames.add("名称");
        columnNames.add("姓名");
        columnNames.add("性别");
        columnNames.add("身份证号");
        columnNames.add("地址");
        writerConfig.setColumnNames(columnNames);
        ArrayList<JSONObject> data = new ArrayList<>();
        for (int i = 0; i < 50; i++) {
            JSONObject json = new JSONObject(true);
            json.put("name", "name" + i);
            json.put("age", i);
            json.put("sex", i % 2 == 0? "男" : "女");
            json.put("card", 555444114);
            json.put("address", null);
            data.add(json);
        }
        writerConfig.setData(data);
        writerConfig.setOutputStream(fileOutputStream);
        AxolotlExcelWriter writer = new AxolotlExcelWriter(writerConfig);
        // TODO 写入
//        writer.write();
        writer.close();

    }

}

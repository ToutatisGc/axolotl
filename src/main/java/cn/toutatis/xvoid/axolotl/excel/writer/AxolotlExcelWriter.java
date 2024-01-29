package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.common.CommonMimeType;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.writer.constant.TemplatePlaceholderPattern;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractInnerStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.CellAddress;
import cn.toutatis.xvoid.axolotl.excel.writer.support.WriteContext;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.axolotl.toolkit.tika.TikaShell;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import com.alibaba.fastjson.JSONObject;
import com.google.common.collect.HashBasedTable;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
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

    public AxolotlExcelWriter(WriterConfig writerConfig) {
        this.writerConfig = writerConfig;
        this.workbook = this.initWorkbook(null);
    }

    public AxolotlExcelWriter(File templateFile, WriterConfig writerConfig) {
        TikaShell.preCheckFileNormalThrowException(templateFile);
        this.writerConfig = writerConfig;
        this.workbook = this.initWorkbook(templateFile);
    }

    private SXSSFWorkbook initWorkbook(File templateFile) {
        SXSSFWorkbook workbook;
        // 读取模板文件内容
        if (templateFile != null){
            LoggerHelper.debug(LOGGER,LoggerHelper.format("正在使用模板文件[%s]作为写入模板",templateFile.getAbsolutePath()));
            AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(templateFile);
            if (!CommonMimeType.OOXML_EXCEL.equals(excelReader.getWorkBookContext().getMimeType())){
                throw new AxolotlWriteException("请使用xlsx文件作为写入模板");
            }
            writeContext.setTemplateReader(excelReader);
            XSSFWorkbook xssfWorkbook = (XSSFWorkbook) writeContext.getTemplateReader().getWorkBookContext().getWorkbook();
            XSSFSheet sheetAt = xssfWorkbook.getSheetAt(0);
            System.err.println(sheetAt);
            workbook = new SXSSFWorkbook(xssfWorkbook);
        }else {
            workbook = new SXSSFWorkbook();
        }
        return workbook;
    }

    public void close() throws IOException {
        workbook.close();
        writerConfig.getOutputStream().close();
    }


    @SneakyThrows
    public void writeToTemplate(int sheetIndex, Map<String,Object> singleMap, List<Object> data) throws IOException {
        SXSSFSheet sheet;
        if (writeContext.isTemplateWrite()){
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()){
                Sheet next = sheetIterator.next();
                System.err.println(next.getSheetName());
                System.err.println(workbook.getSheetIndex(next));
            }
            System.err.println();
            sheet = workbook.getSheetAt(sheetIndex);
            if (sheet != null){
                this.resolveTemplate(sheet);
                if (Validator.objNotNull(singleMap)){
                    // TODO 填充map数据
                }
            }else{
                throw new AxolotlWriteException(LoggerHelper.format("工作表索引[%s]模板中不存在",sheetIndex));
            }
        }else{
            throw new AxolotlWriteException("非模板写入请使用write方法");
        }
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
                        int sheetIndex = workbook.getSheetIndex(sheet);
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
    }

    private void findPlaceholderData(HashBasedTable<Integer, String, CellAddress> data, Pattern pattern, int sheetIndex, CellAddress cellAddress) {
        Matcher matcher = pattern.matcher(cellAddress.getCellValue());
        if (matcher.find()) {
            cellAddress.setPlaceholder(matcher.group());
            data.put(sheetIndex, matcher.group(1), cellAddress);
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

package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.AxolotlTheme;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

/**
 * 文档文件写入器
 * @author Toutatis_Gc
 */
public class AxolotlExcelWriter {

    /**
     * 输出文件
     */
    private final File outputFile;

    private final SXSSFWorkbook workbook;

    public AxolotlExcelWriter(File outputFile) {
        this.outputFile = outputFile;
        workbook = new SXSSFWorkbook();
    }

    public void close() throws IOException {
        workbook.close();
    }


    public void write(WriterConfig writerConfig) throws IOException {
        SXSSFSheet sheet = workbook.createSheet();

//        Object o = data.get(0);
//        System.err.println(o instanceof Map);
//        for (int i = 0; i < data.size(); i++) {
//            SXSSFRow row = sheet.createRow(i);
//            SXSSFCell cell = row.createCell(0);
//            cell.setCellValue(i);
//            SXSSFCell cell1 = row.createCell(1);
//            CellStyle cellStyle = workbook.createCellStyle();
//            cellStyle.setDataFormat(workbook.createDataFormat().getFormat(Time.YMD_HORIZONTAL_FORMAT_REGEX));
//            cell1.setCellValue(LocalDateTime.now());
//            cell1.setCellStyle(cellStyle);
//            SXSSFCell cell2 = row.createCell(2);
//            CellStyle cellStyle2 = workbook.createCellStyle();
//            cellStyle2.setDataFormat(workbook.createDataFormat().getFormat(Time.HMS_COLON_FORMAT_REGEX+"A"));
//            cell2.setCellValue(LocalDateTime.now());
//            cell2.setCellStyle(cellStyle2);
//        }
        AxolotlTheme axolotlTheme = new AxolotlTheme();
        axolotlTheme.setWriterConfig(writerConfig);
        axolotlTheme.renderHeader(sheet);
        workbook.write(new FileOutputStream(outputFile));
    }



    private void exceptionHandler(Exception e){

    }


    public static void main(String[] args) throws IOException {

        AxolotlExcelWriter writer = new AxolotlExcelWriter(new File("D:\\"+IdUtil.randomUUID()+".xlsx"));
        WriterConfig writerConfig = new WriterConfig();
        writerConfig.setTitle("测试生成表标题");
        ArrayList<String> columnNames = new ArrayList<>();
        columnNames.add("名称");
        columnNames.add("性别");
        columnNames.add("身份证号");
        columnNames.add("地址");
        writerConfig.setColumnNames(columnNames);
        writer.write(writerConfig);
//        ArrayList<JSONObject> data = new ArrayList<>();
//        for (int i = 0; i < 50; i++) {
//            JSONObject json = new JSONObject();
//            json.put("id", i);
//            json.put("name", "name" + i);
//            json.put("age", i);
//            json.put("sex", i % 2 == 0? "男" : "女");
//            json.put("address", "address" + i);
//            data.add(json);
//        }
        writer.close();

    }

}

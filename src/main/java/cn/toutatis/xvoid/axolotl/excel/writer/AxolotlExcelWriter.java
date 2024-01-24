package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.excel.writer.style.AbstractInnerStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import com.alibaba.fastjson.JSONObject;
import lombok.SneakyThrows;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

/**
 * 文档文件写入器
 * @author Toutatis_Gc
 */
public class AxolotlExcelWriter {

    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlExcelWriter.class);

    private final SXSSFWorkbook workbook;

    private final WriterConfig writerConfig;

    public AxolotlExcelWriter(WriterConfig writerConfig) {
        this.writerConfig = writerConfig;
        workbook = new SXSSFWorkbook();
    }

    public void close() throws IOException {
        workbook.close();
        writerConfig.getOutputStream().close();
    }


    @SneakyThrows
    public void write() throws IOException {
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



    private void exceptionHandler(Exception e){

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
        writer.write();
        writer.close();

    }

}

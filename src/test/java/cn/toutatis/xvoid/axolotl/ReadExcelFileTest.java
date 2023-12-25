package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.support.WorkBookReaderConfig;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import com.alibaba.fastjson.JSONObject;
import org.junit.Test;

import java.io.File;

public class ReadExcelFileTest {

    @Test
    public void testReadExcelFile() {
        File xlsFile = FileToolkit.getResourceFileAsFile("workbook/1.xls");
        GracefulExcelReader gracefulExcelReader = new GracefulExcelReader(xlsFile);

//        gracefulExcelReader.readData();
        gracefulExcelReader.readSheetData("Sheet1", JSONObject.class);
    }

    @Test
    public void testReadExcelFileWithConfig() {
        File xlsFile = FileToolkit.getResourceFileAsFile("workbook/1.xls");
        WorkBookReaderConfig config = new WorkBookReaderConfig();
        GracefulExcelReader gracefulExcelReader = new GracefulExcelReader(xlsFile,false);
        System.err.println(gracefulExcelReader.readSheetData("Sheet1",Object.class));
    }

}

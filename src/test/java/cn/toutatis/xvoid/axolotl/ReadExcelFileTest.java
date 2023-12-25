package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

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
        GracefulExcelReader gracefulExcelReader = new GracefulExcelReader(xlsFile,false);
        List<Map> mapList = gracefulExcelReader.readSheetData("Sheet1", Map.class);
        System.err.println(JSON.toJSONString(mapList));
    }

}

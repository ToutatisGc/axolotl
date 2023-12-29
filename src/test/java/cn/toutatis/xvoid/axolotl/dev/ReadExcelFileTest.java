package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.excel.GracefulExcelReader;
import cn.toutatis.xvoid.axolotl.entities.IndexPropertyEntity;
import cn.toutatis.xvoid.axolotl.entities.IndexTest;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.serializer.SerializerFeature;
import org.junit.Test;

import java.io.File;
import java.util.List;

public class ReadExcelFileTest {

    @Test
    public void testReadExcelFile() {
        File xlsFile = FileToolkit.getResourceFileAsFile("workbook/1.xls");
        GracefulExcelReader gracefulExcelReader = new GracefulExcelReader(xlsFile);
        List<JSONObject> sheet1 = gracefulExcelReader.readSheetData("Sheet1", JSONObject.class);
        for (JSONObject jsonObject : sheet1) {
            System.err.println(jsonObject);
        }
    }

    @Test
    public void testOneLineExcelFile() {
        File xlsxFile = FileToolkit.getResourceFileAsFile("workbook/单行数据测试.xlsx");
        GracefulExcelReader gracefulExcelReader = new GracefulExcelReader(xlsxFile,true);
        List<IndexTest> mapList = gracefulExcelReader.readSheetData(0, IndexTest.class,0,0);
        System.err.println(new IndexTest());
        for (IndexTest map : mapList) {
            System.err.println(map);
        }
    }

    @Test
    public void testReadExcelFileWithConfig() {
        File xlsFile = FileToolkit.getResourceFileAsFile("workbook/1.xls");
        GracefulExcelReader gracefulExcelReader = new GracefulExcelReader(xlsFile,true);
        List<IndexPropertyEntity> mapList = gracefulExcelReader.readSheetData("Sheet1", IndexPropertyEntity.class);
        System.err.println(JSON.toJSONString(mapList, SerializerFeature.SortField));
    }

}

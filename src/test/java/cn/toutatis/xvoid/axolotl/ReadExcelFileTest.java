package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.entities.IndexPropertyEntity;
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
        gracefulExcelReader.readSheetData("Sheet1", JSONObject.class);
    }

    @Test
    public void testOneLineExcelFile() {
        File xlsxFile = FileToolkit.getResourceFileAsFile("workbook/单行数据测试.xlsx");
        GracefulExcelReader gracefulExcelReader = new GracefulExcelReader(xlsxFile,true);
        List<IndexPropertyEntity> mapList = gracefulExcelReader.readSheetData(0, IndexPropertyEntity.class);
        for (IndexPropertyEntity indexPropertyEntity : mapList) {
            System.err.println(indexPropertyEntity);
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

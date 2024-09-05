package cn.xvoid.axolotl.excel.dev;

import cn.xvoid.axolotl.Axolotls;
import cn.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.xvoid.axolotl.excel.reader.ReadConfigBuilder;
import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.docker.AxolotlCellMapInfo;
import cn.xvoid.toolkit.file.FileToolkit;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

public class ReadTest {

    @Test
    public void testDataValidation() {
        File file = FileToolkit.getResourceFileAsFile("workbook/有效性测试.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<Map> mapList = excelReader.readSheetData(
                new ReadConfigBuilder<>(Map.class)
        );
        System.err.println(mapList);
    }

    @Test
    public void testDataValidation2() {
        File file = FileToolkit.getResourceFileAsFile("workbook/有效性测试.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<Map> mapList = excelReader.readSheetData(
                new ReadConfigBuilder<>(Map.class)
                        .setBooleanReadPolicy(ExcelReadPolicy.MAP_CONVERT_INFO_OBJECT,false)
        );
        System.err.println(mapList);
    }

    @Test
    public void testReadWrapper1() {
        File file = FileToolkit.getResourceFileAsFile("workbook/有效性测试.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<Map<String, AxolotlCellMapInfo>> maps = excelReader.readSheetDataAsMapObject(null);
        System.err.println(maps);
    }

    @Test
    public void testReadWrapper2() {
        File file = FileToolkit.getResourceFileAsFile("workbook/有效性测试.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<Map<String, Object>> maps = excelReader.readSheetDataAsFlatMap(null);
        System.err.println(maps);
    }

}

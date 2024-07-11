package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.ReadConfigBuilder;
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

}

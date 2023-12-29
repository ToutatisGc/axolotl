package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.entities.IndexTest;
import cn.toutatis.xvoid.axolotl.excel.GracefulExcelReader;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import org.junit.Test;

import java.io.File;

public class SimpleTest {

    /**
     * 一般测试
     */
    @Test
    public void testSimple(){
        File file = FileToolkit.getResourceFileAsFile("workbook/单行数据测试.xlsx");
        GracefulExcelReader<IndexTest> excelReader = DocumentLoader.getExcelReader(file, IndexTest.class);

    }

}

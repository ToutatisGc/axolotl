package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.AxolotlDocumentReaders;
import cn.toutatis.xvoid.axolotl.entities.IndexTest;
import cn.toutatis.xvoid.axolotl.excel.AxolotlExcelReader;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import org.junit.Test;

import java.io.File;

public class FileTest {

    @Test
    public void testFile(){
        File file = FileToolkit.getResourceFileAsFile("workbook/2.xlsx");
        // 2.获取文档读取器
        AxolotlExcelReader<IndexTest> excelReader = AxolotlDocumentReaders.getExcelReader(file, IndexTest.class);
    }

}

package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.entities.IndexTest;
import cn.toutatis.xvoid.axolotl.excel.AxolotlExcelReader;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import org.junit.Test;

import java.io.File;

public class SimpleTest {

    /**
     * 一般测试
     */
    @Test
    public void testSimple(){
        // 1.读取Excel文件
        File file = FileToolkit.getResourceFileAsFile("workbook/单行数据测试.xlsx");
        // 2.获取文档读取器
        AxolotlExcelReader<IndexTest> excelReader = AxolotlDocumentReaders.getExcelReader(file, IndexTest.class);
        // 3.读取数据
//        excelReader.readSheetData(0, 0,10 );
    }

}

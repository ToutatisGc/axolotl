package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.entities.IndexTest;
import cn.toutatis.xvoid.axolotl.excel.AxolotlExcelReader;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import org.apache.commons.lang3.time.StopWatch;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.junit.Test;

import java.io.File;
import java.io.IOException;

public class SimpleTest {

    /**
     * 一般测试
     */
    @Test
    public void testSimple(){
        // 1.读取Excel文件
        File file = FileToolkit.getResourceFileAsFile("workbook/大数据量文件.xlsx");
        // 2.获取文档读取器
        AxolotlExcelReader<IndexTest> excelReader = AxolotlDocumentReaders.getExcelReader(file, IndexTest.class);
        // 3.读取数据
//        excelReader.readSheetData(0, 0,10 );
    }

    @Test
    public void testBigFileLoad(){
        File file = FileToolkit.getResourceFileAsFile("workbook/innerBig.xlsx");
        AxolotlExcelReader<IndexTest> excelReader = AxolotlDocumentReaders.getExcelReader(file, IndexTest.class);
    }

    @Test
    public void opcTest(){
        File file = FileToolkit.getResourceFileAsFile("workbook/innerBig.xlsx");
        try {
            StopWatch started = StopWatch.createStarted();
            try (OPCPackage opcPackage = OPCPackage.open(file)) {
                Workbook workbook = XSSFWorkbookFactory.createWorkbook(opcPackage);
                Sheet sheetAt = workbook.getSheetAt(0);
                sheetAt.rowIterator().forEachRemaining(System.out::println);
            }
            started.stop();
            System.err.println(started.getStopTime());
        } catch (InvalidFormatException | IOException e) {
            throw new RuntimeException(e);
        }
    }

}

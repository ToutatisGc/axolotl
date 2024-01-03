package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.AxolotlDocumentReaders;
import cn.toutatis.xvoid.axolotl.entities.OneFieldStringEntity;
import cn.toutatis.xvoid.axolotl.entities.OneFieldStringKeepIntactEntity;
import cn.toutatis.xvoid.axolotl.excel.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

public class SimpleTest {

    /**
     * 空sheet测试
     * 期望读取到0行数据
     */
    @Test
    public void readBlankSheet(){
        File file = FileToolkit.getResourceFileAsFile("workbook/空白表格.xlsx");
        AxolotlExcelReader<Object> excelReader = AxolotlDocumentReaders.getExcelReader(file);
        List<Map> test = excelReader.readSheetData(Map.class, 0);
        Assert.assertEquals(0, test.size());
    }

    @Test
    public void testNamedSheet(){
        File file = FileToolkit.getResourceFileAsFile("workbook/命名空白表格.xlsx");
        AxolotlExcelReader<Object> excelReader = AxolotlDocumentReaders.getExcelReader(file);
        List<Map> mapList = excelReader.readSheetData(Map.class, "表格1");
        Assert.assertEquals(0, mapList.size());
        try {
            excelReader.readSheetData(Map.class, "NON_EXIST");
        }catch (AxolotlExcelReadException e){
            Assert.assertEquals("读取的sheet[NON_EXIST]不存在", e.getMessage());
        }
    }

    @Test
    public void testSingleRowSheet(){
        File file = FileToolkit.getResourceFileAsFile("workbook/单行数据测试.xlsx");
        AxolotlExcelReader<Object> excelReader = AxolotlDocumentReaders.getExcelReader(file);
        List<OneFieldStringEntity> oneFieldStringEntities = excelReader.readSheetData(OneFieldStringEntity.class, 0);
        System.err.println(oneFieldStringEntities);
        List<OneFieldStringKeepIntactEntity> oneFieldStringKeepIntactEntities =
                excelReader.readSheetData(OneFieldStringKeepIntactEntity.class, 0);
        System.err.println(oneFieldStringKeepIntactEntities);
    }

}

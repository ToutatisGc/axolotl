package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.*;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.ReadConfigBuilder;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import com.alibaba.fastjson.JSON;
import com.github.pjfanning.xlsx.StreamingReader;
import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.ValidatorFactory;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class SimpleTest {

    /**
     * 空sheet测试
     * 期望读取到0行数据
     */
    @Test
    public void readBlankSheet(){
        File file = FileToolkit.getResourceFileAsFile("workbook/空白表格.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<Map> test = excelReader.readSheetData(Map.class, 0);
        Assert.assertEquals(0, test.size());
    }

    @Test
    public void testNamedSheet(){
        File file = FileToolkit.getResourceFileAsFile("workbook/命名空白表格.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
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
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<OneFieldStringEntity> oneFieldStringEntities = excelReader.readSheetData(OneFieldStringEntity.class, 0);
        System.err.println(oneFieldStringEntities);
        List<OneFieldStringKeepIntactEntity> oneFieldStringKeepIntactEntities =
                excelReader.readSheetData(OneFieldStringKeepIntactEntity.class, 0);
        System.err.println(oneFieldStringKeepIntactEntities);
    }

    @Test
    public void testOutIndexSheet(){
        File file = FileToolkit.getResourceFileAsFile("workbook/单行数据测试.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<OneFieldStringEntity> oneFieldStringEntities = excelReader.readSheetData(OneFieldStringEntity.class, 10);
        System.err.println(oneFieldStringEntities);
    }

    @Test
    public void testCellMergeSheet(){
        File file = FileToolkit.getResourceFileAsFile("workbook/单元格合并测试.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<OneFieldStringEntity> oneFieldStringEntities = excelReader.readSheetData(OneFieldStringEntity.class, 0);
        System.err.println(oneFieldStringEntities);
    }

    @Test
    public void testMultiData(){
        File file = FileToolkit.getResourceFileAsFile("sec/测试28张表.xls");
        if (file != null && file.exists()){
            AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
            ReadConfigBuilder<DmsRegMonetary> builder = new ReadConfigBuilder<>(DmsRegMonetary.class, true);
            builder.setSheetIndex(0);
            DmsRegMonetary dmsRegMonetary = excelReader.readSheetDataAsObject(builder.build());
            ReadConfigBuilder<DmsRegReceivables> dmsRegReceivablesReadConfigBuilder = new ReadConfigBuilder<>(DmsRegReceivables.class, true);
            dmsRegReceivablesReadConfigBuilder.setSheetIndex(2);
            dmsRegReceivablesReadConfigBuilder.setInitialRowPositionOffset(8);
//            dmsRegReceivablesReadConfigBuilder.setEndIndex();
            List<DmsRegReceivables> dmsRegReceivables = excelReader.readSheetData(dmsRegReceivablesReadConfigBuilder.build());
            for (DmsRegReceivables dmsRegReceivable : dmsRegReceivables) {
                System.err.println(dmsRegReceivable);
            }
            System.err.println(dmsRegReceivables.size());
            try {
                ReadConfigBuilder<DmsMerge> dmsMergeConfig = new ReadConfigBuilder<>(DmsMerge.class);
                dmsMergeConfig
                        .setInitialRowPositionOffset(5)
                        .setEndIndex(28);
                List<DmsMerge> dmsMerges = excelReader.readSheetData(dmsMergeConfig);
                System.err.println(JSON.toJSONString(dmsMerges));
            }catch (Exception e){
                System.err.println(excelReader.getHumanReadablePosition());
                e.printStackTrace();
            }
        }
    }

    @Test
    public void testValidate() {
        try (ValidatorFactory validatorFactory = Validation.buildDefaultValidatorFactory()) {
            Validator validator = validatorFactory.getValidator();
            ValidTestEntity validTestEntity = new ValidTestEntity();
            Set<ConstraintViolation<ValidTestEntity>> validate = validator.validate(validTestEntity);
            for (ConstraintViolation<ValidTestEntity> va : validate) {
                System.err.println(va.getPropertyPath().toString());
                System.err.println(va.getMessage());
            }
        }
    }

    @Test
    public void testHelperConstructor(){

        File file = FileToolkit.getResourceFileAsFile("workbook/单行数据测试.xlsx");
        AxolotlExcelReader<OneFieldStringEntity> excelReader = Axolotls.getExcelReader(file, OneFieldStringEntity.class);
//        List<OneFieldStringEntity> oneFieldStringEntities = excelReader.readSheetData();
//        System.err.println(oneFieldStringEntities);
//        List<OneFieldStringEntity> oneFieldStringEntities1 = excelReader.readSheetData(0, 3);
//        System.err.println(oneFieldStringEntities1);
        List<OneFieldStringEntity> oneFieldStringEntities2 = excelReader.readSheetData(0, 3);
        System.err.println(oneFieldStringEntities2);
    }

    @SneakyThrows
    @Test
    public void itTest(){
        File file = FileToolkit.getResourceFileAsFile("workbook/innerBig.xlsx");
        if (file!= null && file.exists()){
//
            AxolotlExcelReader<OneFieldString3Entity> excelReader = Axolotls.getExcelReader(file, OneFieldString3Entity.class);
            while (excelReader.hasNext()){
                List<OneFieldString3Entity> next = excelReader.next();
                System.err.println(next);
//            System.err.println(next.size());
            }
        }
    }

    @Test
    public void streamTest() throws IOException {
        File file = FileToolkit.getResourceFileAsFile("workbook/innerBig.xlsx");
        if (file!= null && file.exists()){
            InputStream is = new FileInputStream(file);
            Workbook workbook = StreamingReader.builder()
                    .rowCacheSize(1000)
                    .bufferSize(4096)
                    .open(is);
            Sheet sheet = workbook.getSheetAt(0);
            System.err.println(sheet.getLastRowNum());
            int start = 0;
            int end = 1000;
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()){
                Row row = rowIterator.next();
                int rowNum = row.getRowNum();
                if (rowNum >= start && rowNum <= end){
                    Cell cell = row.getCell(3);
                    System.err.println(rowNum+"-"+cell.getStringCellValue());
                }
                if (rowNum > end){
                    break;
                }
            }
            is.close();
        }

    }

    @Test
    public void headerReadTest() {
        File file = FileToolkit.getResourceFileAsFile("workbook/表头读取测试.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<HeaderTestEntity> headerTestEntities = excelReader.readSheetData(HeaderTestEntity.class,0,true,0,-1,1);
        System.err.println(headerTestEntities);
    }

}

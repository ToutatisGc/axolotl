package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.AxolotlDocumentReaders;
import cn.toutatis.xvoid.axolotl.entities.OneFieldStringEntity;
import cn.toutatis.xvoid.axolotl.entities.OneFieldStringKeepIntactEntity;
import cn.toutatis.xvoid.axolotl.entities.ValidTestEntity;
import cn.toutatis.xvoid.axolotl.excel.AxolotlExcelReader;
import cn.toutatis.xvoid.axolotl.excel.ReadConfigBuilder;
import cn.toutatis.xvoid.axolotl.excel.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.ValidatorFactory;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
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

    @Test
    public void testOutIndexSheet(){
        File file = FileToolkit.getResourceFileAsFile("workbook/单行数据测试.xlsx");
        AxolotlExcelReader<Object> excelReader = AxolotlDocumentReaders.getExcelReader(file);
        List<OneFieldStringEntity> oneFieldStringEntities = excelReader.readSheetData(OneFieldStringEntity.class, 10);
        System.err.println(oneFieldStringEntities);
    }

    @Test
    public void testCellMergeSheet(){
        File file = FileToolkit.getResourceFileAsFile("workbook/单元格合并测试.xlsx");
        AxolotlExcelReader<Object> excelReader = AxolotlDocumentReaders.getExcelReader(file);
        List<OneFieldStringEntity> oneFieldStringEntities = excelReader.readSheetData(OneFieldStringEntity.class, 0);
        System.err.println(oneFieldStringEntities);
    }

    @Test
    public void testMultiData(){
        File file = FileToolkit.getResourceFileAsFile("sec/测试28张表.xls");
        if (file != null && file.exists()){
            AxolotlExcelReader<Object> excelReader = AxolotlDocumentReaders.getExcelReader(file);
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

}

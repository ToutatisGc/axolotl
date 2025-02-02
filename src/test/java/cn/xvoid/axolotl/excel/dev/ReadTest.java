package cn.xvoid.axolotl.excel.dev;

import cn.xvoid.axolotl.AxolotlFaster;
import cn.xvoid.axolotl.Axolotls;
import cn.xvoid.axolotl.excel.entities.reader.DmsMerge;
import cn.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.xvoid.axolotl.excel.reader.ReadConfigBuilder;
import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.hooks.BatchReadTask;
import cn.xvoid.axolotl.excel.reader.hooks.ReadProgressHook;
import cn.xvoid.axolotl.excel.reader.support.docker.AxolotlCellMapInfo;
import cn.xvoid.toolkit.file.FileToolkit;
import cn.xvoid.toolkit.number.Calculator;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

public class ReadTest {

    @Test
    public void testDataValidation() {
        File file = FileToolkit.getResourceFileAsFile("sec/issue2.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<Map> mapList = excelReader.readSheetData(
                new ReadConfigBuilder<>(Map.class).build(),
                (current, total) -> System.err.println("当前进度：" + Calculator.evaluateAsPlainText("%d/%d*%d".formatted(current, total, 100)) + "%")
        );
        System.err.println(mapList);
    }

    @Test
    public void testnum(){
        System.err.println(1/10);
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

    @Test
    public void testReadBatchReadData() {
        AxolotlExcelReader<DmsMerge> excelReader = Axolotls.getExcelReader(new File("C:\\Users\\Administrator\\Desktop\\ceshi.xlsx"), DmsMerge.class);
       /* ReaderConfig<DmsMerge> readerConfig = new ReaderConfig<>();
        readerConfig.setCastClass(DmsMerge.class);
        readerConfig.setStartIndex(40);
        readerConfig.setSheetColumnEffectiveRange(0,999);

        excelReader.batchReadData(10, readerConfig, new BatchReadTask<DmsMerge>() {
            @Override
            public void execute(List<DmsMerge> data) {
                System.out.println(data);
            }
        }, new ReadProgressHook() {
            @Override
            public void onReadProgress(int current, int total) {
                System.out.println("["+current+"]-["+total+"]");
            }
        });*/

        AxolotlFaster.batchReadData(excelReader, 25, new BatchReadTask<DmsMerge>() {
            @Override
            public void execute(List<DmsMerge> data) {
                System.out.println(data);
            }
        },  new ReadProgressHook() {
            @Override
            public void onReadProgress(int current, int total) {
                System.out.println("["+current+"]-["+total+"]");
            }
        }, DmsMerge.class,0,null,0,7);

    }

}

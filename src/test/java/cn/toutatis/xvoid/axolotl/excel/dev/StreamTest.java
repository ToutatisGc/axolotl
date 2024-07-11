package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.toutatis.xvoid.axolotl.AxolotlFaster;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.reader.OneFieldString3Entity;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlStreamExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.reader.support.stream.AxolotlExcelStream;
import cn.toutatis.xvoid.axolotl.excel.reader.support.stream.ReadBatchTask;
import cn.xvoid.toolkit.file.FileToolkit;
import org.junit.Test;

import java.io.File;
import java.util.List;
import java.util.Map;

public class StreamTest {

    @Test
    public void streamTest1(){
        File file = FileToolkit.getResourceFileAsFile("workbook/innerBig.xlsx");
        if (file!= null && file.exists()){
            AxolotlStreamExcelReader<Object> streamExcelReader = Axolotls.getStreamExcelReader(file);
            int recordRowNumber = streamExcelReader.getRecordRowNumber();
            System.err.println(recordRowNumber);
            ReaderConfig<OneFieldString3Entity> readerConfig = new ReaderConfig<>(OneFieldString3Entity.class);
            AxolotlExcelStream<OneFieldString3Entity> dataIterator = streamExcelReader.dataIterator(readerConfig);
            int idx = 0;
            while (dataIterator.hasNext()){
              /*  OneFieldString3Entity entity = dataIterator.next();
                System.out.println(idx+"="+entity);*/
                idx++;
            }
        }
    }


    @Test
    public void streamTest2(){
        File file = new File("D:\\46399f1b-5e02-46b7-bf72-886fab894448.xlsx");
        AxolotlExcelStream<OneFieldString3Entity> stream = AxolotlFaster.readSheetUseStream(AxolotlFaster.getStreamExcelReader(file),
                OneFieldString3Entity.class,
                0,
                0,
                -1,
                0,
                -1,
                false
                );
        stream.readDataBatch(2, new ReadBatchTask<OneFieldString3Entity>() {
            @Override
            public void execute(List<OneFieldString3Entity> data) {
                System.out.println(data);
            }
        });

    }

    @Test
    public void streamTest3(){
        File file = new File("D:\\46399f1b-5e02-46b7-bf72-886fab894448.xlsx");
        List<OneFieldString3Entity> list = AxolotlFaster.readSheetAsList(AxolotlFaster.getExcelReader(file),
                OneFieldString3Entity.class,
                0,
                0,
                -1,
                0,
                -1,
                false
        );
        /*stream.readDataBatch(1, new ReadBatchTask<OneFieldString3Entity>() {
            @Override
            public void execute(List<OneFieldString3Entity> data) {
                System.out.println(data);
            }
        });
*/
    }

}

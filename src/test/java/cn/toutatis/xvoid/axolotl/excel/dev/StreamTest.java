package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.OneFieldString3Entity;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlStreamExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.reader.support.stream.AxolotlExcelStream;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import org.junit.Test;

import java.io.File;

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
                OneFieldString3Entity entity = dataIterator.next();
                System.out.println(idx+"="+entity);
                idx++;
            }
        }
    }

}

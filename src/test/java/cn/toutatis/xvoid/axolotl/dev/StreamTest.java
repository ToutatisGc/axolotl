package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlStreamExcelReader;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import org.junit.Test;

import java.io.File;

public class StreamTest {

    @Test
    public void streamTest1(){
        File file = FileToolkit.getResourceFileAsFile("workbook/innerBig.xlsx");
        if (file!= null && file.exists()){
            AxolotlStreamExcelReader<Object> streamExcelReader = Axolotls.getStreamExcelReader(file);
        }

    }

}

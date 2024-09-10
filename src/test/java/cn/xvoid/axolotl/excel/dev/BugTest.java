package cn.xvoid.axolotl.excel.dev;

import cn.xvoid.axolotl.Axolotls;
import cn.xvoid.axolotl.excel.entities.reader.Members;
import cn.xvoid.axolotl.excel.reader.AxolotlExcelReader;
import cn.xvoid.toolkit.file.FileToolkit;
import org.junit.Test;

import java.io.File;
import java.util.List;

public class BugTest {

    @Test
    public void testReadBug(){
        File file = FileToolkit.getResourceFileAsFile("sec/issue1.xlsx");
        AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
        List<Members> members = excelReader.readSheetData(Members.class, 0);
        System.err.println(members);
    }

}

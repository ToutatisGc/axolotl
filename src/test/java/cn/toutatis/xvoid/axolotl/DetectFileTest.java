package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.constant.CommonMimeType;
import cn.toutatis.xvoid.axolotl.support.DetectResult;
import cn.toutatis.xvoid.axolotl.support.TikaShell;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import org.apache.tika.exception.TikaException;
import org.apache.tika.mime.MimeType;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.IOException;

public class DetectFileTest {

    /**
     * 测试识别文件类型
     * @throws IOException IOException
     * @throws TikaException TikaException
     */
    @Test
    public void matchExcelFile() throws IOException, TikaException {
        File xlsFile = FileToolkit.getResourceFileAsFile("workbook/1.xls");
        File xlsxFile = FileToolkit.getResourceFileAsFile("workbook/2.xlsx");

        MimeType xlsMimeType = TikaShell.getMimeType(xlsFile);
        Assert.assertEquals(xlsMimeType, TikaShell.MS_EXCEL);

        MimeType xlsxMimeType = TikaShell.getMimeType(xlsxFile);
        Assert.assertEquals(xlsxMimeType, TikaShell.OOXML_EXCEL);
    }

    @Test
    public void detectExcelFile() {
        File xlsFile = FileToolkit.getResourceFileAsFile("workbook/1.xls");
        DetectResult detect1 = TikaShell.detect(xlsFile, CommonMimeType.MS_EXCEL);
        Assert.assertTrue(detect1.isDetect());
        DetectResult detect2 = TikaShell.detect(xlsFile, CommonMimeType.OOXML_EXCEL);
        Assert.assertFalse(detect2.isDetect());
        File xlsxFile = new File(FileToolkit.getResourceFile("workbook/2.xlsx").getFile());
        DetectResult detect3 = TikaShell.detect(xlsxFile, CommonMimeType.MS_EXCEL);
        Assert.assertFalse(detect3.isDetect());
        DetectResult detect4 = TikaShell.detect(xlsxFile, CommonMimeType.OOXML_EXCEL);
        Assert.assertTrue(detect4.isDetect());
    }

}

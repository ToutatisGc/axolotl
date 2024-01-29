package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.WriterConfig;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WriteTest {

    @Test
    public void findTemplateKey() {
        String input = "This is a ${test} string with #{multiple} placeholders.";

        Pattern pattern = Pattern.compile("\\$\\{([^}]*)\\}");
        Matcher matcher = pattern.matcher(input);
        boolean b = matcher.find();
        Assert.assertTrue(b);
        Assert.assertEquals("test", matcher.group(1));

        Pattern pattern1 = Pattern.compile("#\\{([^}]*)\\}");
        Matcher matcher1 = pattern1.matcher(input);
        boolean b1 = matcher1.find();
        Assert.assertTrue(b1);
        Assert.assertEquals("multiple", matcher1.group(1));
    }

    @Test
    public void testWritePlaceholders() throws Exception {
        File file = FileToolkit.getResourceFileAsFile("workbook/读取占位符测试.xlsx");
        WriterConfig writerConfig = new WriterConfig();
        AxolotlExcelWriter axolotlExcelWriter = new AxolotlExcelWriter(file, null);
        axolotlExcelWriter.writeToTemplate(0, null, null);
    }

}

package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.WriterConfig;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import com.alibaba.fastjson.JSONObject;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Map;
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
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        writerConfig.setOutputStream(fileOutputStream);
        AxolotlExcelWriter axolotlExcelWriter = new AxolotlExcelWriter(file, writerConfig);
        Map<String, String> map = Map.of("fix2", "测试内容2", "fix1", new SimpleDateFormat(Time.YMD_HORIZONTAL_FORMAT_REGEX).format(Time.getCurrentMillis()));
        ArrayList<JSONObject> data = new ArrayList<>();
        for (int i = 0; i < 50; i++) {
            JSONObject json = new JSONObject(true);
            json.put("name", "name" + i);
            json.put("age", i);
            json.put("sex", i % 2 == 0? "男" : "女");
            json.put("card", 555444114);
            json.put("address", null);
            data.add(json);
        }
        axolotlExcelWriter.writeToTemplate(0, map, data);
    }

}

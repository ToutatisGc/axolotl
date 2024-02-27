package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.RandomUtil;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.WriterConfig;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.lang3.RandomStringUtils;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
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
        File file = FileToolkit.getResourceFileAsFile("workbook/write/读取占位符测试.xlsx");
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
        ArrayList<DmsRegReceivables> LIST = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            DmsRegReceivables dmsRegReceivables = new DmsRegReceivables();
            dmsRegReceivables.setReceivablesDebtor("测试"+i);
            dmsRegReceivables.setReceivablesApprover(RandomStringUtils.randomAlphabetic(32));
            dmsRegReceivables.setReceivablesExpirationLocalDateTime(LocalDateTime.now());
            dmsRegReceivables.setReceivablesExpirationDate(new Date());
            dmsRegReceivables.setReceivablesVerifyBigDecimal(new BigDecimal("3.55"));
            LIST.add(dmsRegReceivables);
        }

        axolotlExcelWriter.writeToTemplate(0, map, LIST);
        axolotlExcelWriter.close();
    }

    @Test
    public void testShift() throws IOException {
        File file = FileToolkit.getResourceFileAsFile("workbook/write/漂移写入测试.xlsx");
        WriterConfig writerConfig = new WriterConfig();
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        writerConfig.setOutputStream(fileOutputStream);
        try (AxolotlExcelWriter axolotlExcelWriter = Axolotls.getTemplateExcelWriter(file, writerConfig)) {
            Map<String, Object> map = Map.of("name", "Toutatis","nation","汉");
            axolotlExcelWriter.writeToTemplate(0, map, null);
            ArrayList<JSONObject> datas = new ArrayList<>();
            for (int i = 0; i < 3; i++) {
                JSONObject sch = new JSONObject();
                sch.put("schoolName","北京-"+RandomStringUtils.randomAlphabetic(16));
                sch.put("schoolYears", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP));
                sch.put("graduate", true);
                datas.add(sch);
            }
            axolotlExcelWriter.writeToTemplate(0, Map.of("age",50), datas);
        }
    }

}

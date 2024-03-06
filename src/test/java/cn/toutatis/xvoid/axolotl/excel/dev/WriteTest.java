package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.RandomUtil;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.reader.DmsRegReceivables;
import cn.toutatis.xvoid.axolotl.excel.entities.reader.SunUser;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.AnnoEntity;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlTemplateExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.TemplateWriteConfig;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.file.FileToolkit;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.util.*;
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
        TemplateWriteConfig commonWriteConfig = new TemplateWriteConfig();
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        commonWriteConfig.setOutputStream(fileOutputStream);
        AxolotlTemplateExcelWriter axolotlAutoExcelWriter = Axolotls.getTemplateExcelWriter(file, commonWriteConfig);
        Map<String, String> map = new HashMap<>();
        map.put("fix2", "测试内容2");
        map.put( "fix1", new SimpleDateFormat(Time.YMD_HORIZONTAL_FORMAT_REGEX).format(Time.getCurrentMillis()));
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

        axolotlAutoExcelWriter.write(map, LIST);
        axolotlAutoExcelWriter.close();
    }

    @Test
    public void testShift() throws IOException {
        File file = FileToolkit.getResourceFileAsFile("workbook/write/漂移写入测试.xlsx");
        TemplateWriteConfig commonWriteConfig = new TemplateWriteConfig();
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        commonWriteConfig.setOutputStream(fileOutputStream);
        try (AxolotlTemplateExcelWriter axolotlAutoExcelWriter = Axolotls.getTemplateExcelWriter(file, commonWriteConfig)) {
            ArrayList<JSONObject> datas1 = new ArrayList<>();
            for (int i = 0; i < 3; i++) {
                JSONObject sch = new JSONObject();
                sch.put("workPlace","北京-"+RandomStringUtils.randomAlphabetic(16));
                sch.put("workYears", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP));
                sch.put("salary", true);
                datas1.add(sch);
            }
            axolotlAutoExcelWriter.write(null, datas1);
            Map<String, Object> map = new HashMap<>();
            map.put("name", "Toutatis");
            map.put("nation","汉");
            axolotlAutoExcelWriter.write(map, null);
            ArrayList<JSONObject> datas = new ArrayList<>();
            for (int i = 0; i < 3; i++) {
                JSONObject sch = new JSONObject();
                sch.put("schoolName","北京-"+RandomStringUtils.randomAlphabetic(16));
                sch.put("schoolYears", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP));
                sch.put("graduate", true);
                datas.add(sch);
            }
            Map<String, Object> map2 = new HashMap<>();
            map2.put("age",50);
            axolotlAutoExcelWriter.write(map2, datas);
            datas.clear();
            for (int i = 0; i < 5; i++) {
                JSONObject sch = new JSONObject();
                sch.put("schoolName","北京-"+RandomStringUtils.randomAlphabetic(16));
                sch.put("schoolYears", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP));
                sch.put("graduate", true);
                datas.add(sch);
            }
            axolotlAutoExcelWriter.write(null, datas);
        }
    }
    @Test
    public void test() throws FileNotFoundException {
        File file = FileToolkit.getResourceFileAsFile("workbook/write/sunUser.xlsx");
        TemplateWriteConfig commonWriteConfig = new TemplateWriteConfig();
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        commonWriteConfig.setOutputStream(fileOutputStream);
// 创建写入器
        try (AxolotlTemplateExcelWriter axolotlAutoExcelWriter = new AxolotlTemplateExcelWriter(file, commonWriteConfig)) {
            List list = new ArrayList();
            for (int i = 0; i < 100; i++) {
                SunUser sunUser = new SunUser();
                for (Field declaredField : SunUser.class.getDeclaredFields()) {
                    ReflectToolkit.setObjectField(sunUser, declaredField, "11"+i);
                }
                list.add(sunUser);
            }
            Map<String, String> map = new HashMap<>();
            map.put("name","测试姓名");
            axolotlAutoExcelWriter.write(map,list);
        }

    }

    @Test
    public void testAuto1() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(new File("D:\\" + IdUtil.randomUUID() + ".xlsx"));
        AutoWriteConfig commonWriteConfig = new AutoWriteConfig();
        commonWriteConfig.setOutputStream(fileOutputStream);
        commonWriteConfig.setTitle("测试生成表标题");
        ArrayList<String> columnNames = new ArrayList<>();
        columnNames.add("名称");
        columnNames.add("姓名");
        columnNames.add("性别");
        columnNames.add("身份证号");
        columnNames.add("地址");
//        commonWriteConfig.setColumnNames(columnNames);
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
        AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(commonWriteConfig);
        autoExcelWriter.write(null,data);
        autoExcelWriter.close();

    }

    @Test
    public void testAnno(){
        List<AnnoEntity> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            AnnoEntity annoEntity = new AnnoEntity();
            annoEntity.setName("name"+i);
            annoEntity.setAddress("address"+i);
            list.add(annoEntity);
        }
        AutoWriteConfig autoWriteConfig = new AutoWriteConfig();
        AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(autoWriteConfig);
        autoExcelWriter.write(list);
    }

    @Test
    public void generateColorCard() throws IOException {
        SXSSFWorkbook workbook = new SXSSFWorkbook();
        SXSSFSheet sheet = workbook.createSheet("内置色卡");
        for (int i = 0; i < FillPatternType.values().length; i++) {
            FillPatternType patternType = FillPatternType.values()[i];
            SXSSFRow row = sheet.createRow(i);
            SXSSFCell cell = row.createCell(0);
            cell.setCellValue(patternType.name());
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillPattern(patternType);
            cellStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
            cell.setCellStyle(cellStyle);
        }
//        for (int i = 0; i < IndexedColors.values().length; i++) {
//            SXSSFRow row = sheet.createRow(i);
//            SXSSFCell cell = row.createCell(0);
//            IndexedColors color = IndexedColors.values()[i];
//            cell.setCellValue(color.name());
//            CellStyle cellStyle = workbook.createCellStyle();
//            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//            cellStyle.setFillForegroundColor(color.getIndex());
//            cell.setCellStyle(cellStyle);
//        }
        workbook.write(new FileOutputStream("D:\\COLOR-CARD.xlsx"));
        workbook.close();
    }
}

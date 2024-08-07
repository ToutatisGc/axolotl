package cn.xvoid.axolotl.excel.dev;

import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.RandomUtil;
import cn.xvoid.axolotl.AxolotlFaster;
import cn.xvoid.axolotl.Axolotls;
import cn.xvoid.axolotl.excel.entities.reader.DmsRegReceivables;
import cn.xvoid.axolotl.excel.entities.reader.SunUser;
import cn.xvoid.axolotl.excel.entities.writer.MpOrgDataIssueNew;
import cn.xvoid.axolotl.excel.writer.AxolotlTemplateExcelWriter;
import cn.xvoid.axolotl.excel.writer.TemplateWriteConfig;
import cn.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.xvoid.toolkit.clazz.ReflectToolkit;
import cn.xvoid.toolkit.constant.Time;
import cn.xvoid.toolkit.file.FileToolkit;
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

import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
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

        axolotlAutoExcelWriter.write(map, LIST);
        axolotlAutoExcelWriter.close();
    }

    @Test
    public void testShift() throws IOException {
        File file = FileToolkit.getResourceFileAsFile("workbook/write/漂移写入测试.xlsx");
        TemplateWriteConfig commonWriteConfig = new TemplateWriteConfig();
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        commonWriteConfig.setOutputStream(fileOutputStream);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_PLACEHOLDER_FILL_DEFAULT,false);
        //commonWriteConfig.setWritePolicy(ExcelWritePolicy.TEMPLATE_NON_TEMPLATE_CELL_FILL,false);
        try (AxolotlTemplateExcelWriter axolotlAutoExcelWriter = Axolotls.getTemplateExcelWriter(file, commonWriteConfig)) {
            ArrayList<JSONObject> datas1 = new ArrayList<>();
            for (int i = 0; i < 2; i++) {
                JSONObject sch = new JSONObject();
                sch.put("workPlace","北京-"+RandomStringUtils.randomAlphabetic(16));
                sch.put("workYears1", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP));
                sch.put("salary1", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP).multiply(new BigDecimal("1000")));
                datas1.add(sch);
            }
            axolotlAutoExcelWriter.write(null, datas1);
            /*ArrayList<JSONObject> datas2 = new ArrayList<>();
            for (int i = 0; i < 2; i++) {
                JSONObject sch = new JSONObject();
                sch.put("workPlace","北京-"+RandomStringUtils.randomAlphabetic(16));
                sch.put("workYears", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP));
                sch.put("salary", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP).multiply(new BigDecimal("1000")));
                datas2.add(sch);
            }
            axolotlAutoExcelWriter.write(null, datas2);*/
            /*Map<String, Object> map = Map.of("name", "Toutatis","nation","汉");
            axolotlAutoExcelWriter.write(map, null);
            ArrayList<JSONObject> datas = new ArrayList<>();
            for (int i = 0; i < 3; i++) {
                JSONObject sch = new JSONObject();
                sch.put("schoolName","北京-"+RandomStringUtils.randomAlphabetic(16));
                sch.put("schoolYears", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP));
                sch.put("graduate", true);
                datas.add(sch);
            }
            axolotlAutoExcelWriter.write(Map.of("age",50), datas);
            datas.clear();
            for (int i = 0; i < 5; i++) {
                JSONObject sch = new JSONObject();
                sch.put("schoolName","北京-"+RandomStringUtils.randomAlphabetic(16));
                sch.put("schoolYears", RandomUtil.randomBigDecimal(BigDecimal.ZERO, BigDecimal.TEN).setScale(0, RoundingMode.HALF_UP));
                sch.put("graduate", true);
                datas.add(sch);
            }
            axolotlAutoExcelWriter.write(null, datas);*/
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
    public void testBlankFill() throws FileNotFoundException {
        File file = FileToolkit.getResourceFileAsFile("workbook/write/空白占位符测试.xlsx");
        TemplateWriteConfig commonWriteConfig = new TemplateWriteConfig();
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        commonWriteConfig.setOutputStream(fileOutputStream);
        // 创建写入器
        try (AxolotlTemplateExcelWriter excelWriter = new AxolotlTemplateExcelWriter(file, commonWriteConfig)) {
            List list = new ArrayList();
            for (int i = 0; i < 2; i++) {
                JSONObject jsonObject = new JSONObject(true);
                jsonObject.put("f1","tt");
                list.add(jsonObject);
            }
            excelWriter.write(null,list);
        }
    }

//    @Test
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

    @Test
    public void test1() throws FileNotFoundException {
          File file = FileToolkit.getResourceFileAsFile("workbook/write/dataScheduleOther.xlsx");
//        File file = new File("D:\\dataScheduleOther.xlsx");
        TemplateWriteConfig commonWriteConfig = new TemplateWriteConfig();
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        commonWriteConfig.setOutputStream(fileOutputStream);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.SIMPLE_USE_GETTER_METHOD,true);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.SIMPLE_USE_DICT_CODE_TRANSFER,true);
        commonWriteConfig.setDict("regionStatus",Map.of("ST_001","正常"));
// 创建写入器
        try (AxolotlTemplateExcelWriter axolotlAutoExcelWriter = new AxolotlTemplateExcelWriter(file, commonWriteConfig)) {
            List list = new ArrayList();

            for (int i = 0; i < 20; i++) {
                MpOrgDataIssueNew mpOrgDataIssueNew = new MpOrgDataIssueNew();
                list.add(mpOrgDataIssueNew);
            }
            Map<String, String> map = new HashMap<>();
            map.put("fileName","测试文件名");
            map.put("bankName","山西省");
//            map.put("dataIssue","2024-02");
            map.put("operationTime", LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
            axolotlAutoExcelWriter.write(map,list);
        }

    }

    @Test
    public void test11() throws FileNotFoundException {
        File tem = FileToolkit.getResourceFileAsFile("workbook/write/dataScheduleOther.xlsx");
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        List list = new ArrayList();

        for (int i = 0; i < 20; i++) {
            MpOrgDataIssueNew mpOrgDataIssueNew = new MpOrgDataIssueNew();
            list.add(mpOrgDataIssueNew);
        }
        Map<String, String> map = new HashMap<>();
        map.put("fileName","测试文件名");
        map.put("bankName","山西省");
//            map.put("dataIssue","2024-02");
        map.put("operationTime", LocalDate.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd")));
       // AxolotlFaster.writeToTemplate(tem,fileOutputStream,map,list);
        Map<String, Object> dict = new HashMap<>();
        dict.put("regionStatus",Map.of("ST_001","正常"));
        HashMap<String, Object> dict1 = new HashMap<>();
        dict1.put("regionStatus",Map.of("ST_001","正常1"));
        AxolotlFaster.templateWriteToExcelMultiSheet(tem, fileOutputStream,
                AxolotlFaster.buildTemplateWriteSheetInfo(0,map,list,true,true,dict),
                AxolotlFaster.buildTemplateWriteSheetInfo(0,map,list,true,true,dict1),
                AxolotlFaster.buildTemplateWriteSheetInfo(1,map,list,true,true,dict)
                );
    }
}

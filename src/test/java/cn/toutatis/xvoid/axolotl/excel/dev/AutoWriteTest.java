package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.RandomUtil;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.AnnoEntity;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.StockEntity;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.components.configuration.AxolotlCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.configuration.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.AxolotlSelectBox;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.Header;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.AxolotlDefaultStyleConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.standard.AxolotlMidnightTheme;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.configurable.AxolotlConfigurableTheme;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.ExcelWriteThemes;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.*;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public class AutoWriteTest {


    @Test
    public void testAuto1() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        AutoWriteConfig commonWriteConfig = new AutoWriteConfig();
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,true);
        commonWriteConfig.setTitle("测试表");
        commonWriteConfig.setOutputStream(fileOutputStream);
        List<Header> headers = new ArrayList<>();
        headers.add(new Header("名称",List.of(new Header("姓名"),new Header("花名"))));
        headers.add(new Header("期限", List.of(new Header("年"), new Header("月"))));
        AxolotlCellStyle axolotlCellStyle1 = new AxolotlCellStyle();
        axolotlCellStyle1.setForegroundColor(new AxolotlColor(0,231,185));
        Header header1 = new Header("金额");
        header1.setColumnWidth(10000);
        Header header = new Header("账面数",
                List.of(new Header("经济",
                        List.of(new Header("数量"), new Header("金额"))), new Header("数量"), header1));
        header.setAxolotlCellStyle(axolotlCellStyle1);
        headers.add(header);
        Header remark = new Header("备注");
//        remark.setFieldName("remark");
        remark.setColumnWidth(10000);
        AxolotlCellStyle axolotlCellStyle = new AxolotlCellStyle();
        axolotlCellStyle.setForegroundColor(new AxolotlColor(155,231,185));
        remark.setAxolotlCellStyle(axolotlCellStyle);
        headers.add(remark);
        ArrayList<JSONObject> data = new ArrayList<>();
        for (int i = 0; i < 50; i++) {
            JSONObject json = new JSONObject(true);
            json.put("remark", "name" + i);
            json.put("age", i);
            json.put("sex", i % 2 == 0? "男" : "女");
            json.put("card", 555444114);
            json.put("address", null);
            data.add(json);
        }
        AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(commonWriteConfig);
        autoExcelWriter.write(headers,data);
        autoExcelWriter.close();
    }

    @Test
    public void testAuto2() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        AutoWriteConfig commonWriteConfig = new AutoWriteConfig();
        commonWriteConfig.setThemeStyleRender(ExcelWriteThemes.ADMINISTRATION_RED);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING,true);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH,true);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,true);
        commonWriteConfig.setTitle("股票测试表");
        commonWriteConfig.setFontName("仿宋");
        commonWriteConfig.setBlankValue("-");
        commonWriteConfig.setOutputStream(fileOutputStream);
        List<Header> headers = new ArrayList<>();
        headers.add(new Header("代码","code"));
        headers.add(new Header("简称","intro"));
        headers.add(new Header("最新日期","localDateTimeStr"));
        headers.add(new Header("最新收盘价（元）",StockEntity::getClosingPrice));
        headers.add(new Header("涨跌幅（%）",StockEntity::getPriceLimit));
        headers.add(new Header("总市值（亿元）",StockEntity::getTotalValue));
        headers.add(new Header("流通市值（亿元）",StockEntity::getCirculationMarketValue));
        List<Header> subHeader1 = List.of(new Header("TTM"), new Header("15E"),new Header("16E"));
        headers.add(new Header("每股收益", subHeader1));
        headers.add(new Header("市盈率PE", subHeader1));
        headers.add(new Header("市净率PB（LF）"));
        headers.add(new Header("市销率PS（TTM）","pts"));
        ArrayList<StockEntity> data = new ArrayList<>();
        for (int i = 0; i < 50; i++) {
            StockEntity stockEntity = new StockEntity();
            stockEntity.setCode(RandomStringUtils.randomNumeric(8));
            StringBuilder sb = new StringBuilder();
            for (int i1 = 0; i1 < 10; i1++) {
                char c = RandomUtil.randomChinese();
                sb.append(c);
            }
            stockEntity.setIntro(sb.toString());
            stockEntity.setPts(Double.parseDouble(RandomStringUtils.randomNumeric(2)));
            data.add(stockEntity);
        }
        AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(commonWriteConfig);
        autoExcelWriter.write(headers,data);
        autoExcelWriter.close();

    }

    @Test
    public void autoTest3() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        AutoWriteConfig config = new AutoWriteConfig();
        config.setTitle("测试1");
        config.setThemeStyleRender(AxolotlMidnightTheme.class);
        config.setFontName("微软雅黑");
        config.setWritePolicy(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH,true);
        config.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,true);
        config.setBlankValue("-");
        config.setOutputStream(fileOutputStream);
        List<Header> headers = new ArrayList<>();
        headers.add(new Header("Capacity\nTON"));
        headers.add(new Header("Model",new Header("Single\nSpeed"),new Header("Dual\nSpeed")));
        headers.add(new Header("I-Beam",new Header("(mm)"),new Header("(inch)")));
        headers.add(new Header("Travel Speed",new Header("50HZ"),new Header("60HZ")));
        headers.add(new Header("Motor(KW)",new Header("Single",true),new Header("Dual",true)));
        config.addSpecialRowHeight(1,40);
        ArrayList<JSONObject> data = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            JSONObject jsonObject = new JSONObject(true);
            jsonObject.put("order",i);
            jsonObject.put("model","CHM-00"+i);
            jsonObject.put("dual speed","CD-00"+i);
            jsonObject.put("mm",RandomUtil.randomInt(0,200)+"-"+RandomUtil.randomInt(200,500));
            jsonObject.put("inch",RandomUtil.randomInt(0,10)+"°-"+RandomUtil.randomInt(10,20)+"°");
            jsonObject.put("50HZ",RandomUtil.randomDouble(0,20));
            jsonObject.put("60HZ",RandomUtil.randomDouble(0,20));
            jsonObject.put("Single",RandomUtil.randomDouble(0,20));
            jsonObject.put("Dual",RandomUtil.randomDouble(0,20));
            data.add(jsonObject);
        }
        AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(config);
        autoExcelWriter.write(headers,data);
        autoExcelWriter.close();

    }

    @Test
    public void testAnno() throws IOException {
        List<AnnoEntity> list = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            AnnoEntity annoEntity = new AnnoEntity();
            annoEntity.setName("name"+i);
            annoEntity.setAddress("address"+i);
            list.add(annoEntity);
        }
        AutoWriteConfig autoWriteConfig = new AutoWriteConfig();
        autoWriteConfig.setClassMetaData(list);
        autoWriteConfig.setOutputStream(new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx"));
        AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(autoWriteConfig);
        autoExcelWriter.write(list);
        autoExcelWriter.close();
    }

    @Test
    public void testCalculateHeaders(){
        List<Header> headers = new ArrayList<>();
        headers.add(new Header("名称"));
        headers.add(new Header("期限", List.of(new Header("年"),new Header("月"))));
        headers.add(new Header("账面数", List.of(new Header("经济",List.of(new Header("数量"),new Header("金额"))),new Header("非经济",List.of(new Header("数量"),new Header("金额"))))));
        int maxDepth = ExcelToolkit.getMaxDepth(headers, 0);
        Assert.assertEquals(3,maxDepth);
        for (Header header : headers) {
            System.err.println(header.countOrlopCellNumber());
        }
    }

    @Test
    public void testAnnoGetHeaderList(){
        List<Header> headers = ExcelToolkit.getHeaderList(AnnoEntity.class);
        for (Header header : headers) {
            System.err.println(header.getName());
        }
    }

    @Test
    public void testAuto4() throws IOException {
        for (ExcelWriteThemes theme : ExcelWriteThemes.values()) {
            FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
            AutoWriteConfig autoWriteConfig = new AutoWriteConfig();
            autoWriteConfig.setThemeStyleRender(theme);
            autoWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH,true);
            autoWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,true);
            autoWriteConfig.setTitle("股票测试表");
            autoWriteConfig.setBlankValue("-");
            autoWriteConfig.setOutputStream(fileOutputStream);
            List<Header> headers = new ArrayList<>();
            headers.add(new Header("代码","code"));
            headers.add(new Header("简称","intro"));
            headers.add(new Header("最新日期","localDateTimeStr"));
            headers.add(new Header("最新收盘价（元）",StockEntity::getClosingPrice));
            headers.add(new Header("涨跌幅（%）",StockEntity::getPriceLimit));
            headers.add(new Header("总市值（亿元）",StockEntity::getTotalValue));
            headers.add(new Header("流通市值（亿元）",StockEntity::getCirculationMarketValue));
            List<Header> subHeader1 = List.of(new Header("TTM"), new Header("15E"),new Header("16E"));
            headers.add(new Header("每股收益", subHeader1));
            headers.add(new Header("市盈率PE", subHeader1));
            headers.add(new Header("市净率PB（LF）"));
            headers.add(new Header("市销率PS（TTM）","pts"));
            ArrayList<StockEntity> data = new ArrayList<>();
            for (int i = 0; i < 50; i++) {
                StockEntity stockEntity = new StockEntity();
                stockEntity.setCode(RandomStringUtils.randomNumeric(8));
                StringBuilder sb = new StringBuilder();
                for (int i1 = 0; i1 < 10; i1++) {
                    char c = RandomUtil.randomChinese();
                    sb.append(c);
                }
                stockEntity.setIntro(sb.toString());
                stockEntity.setPts(Double.parseDouble(RandomStringUtils.randomNumeric(2)));
                data.add(stockEntity);
            }
            AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(autoWriteConfig);
            autoExcelWriter.write(headers,data);
            autoExcelWriter.close();
        }
    }

    @Test
    public void testAuto5() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        AutoWriteConfig commonWriteConfig = new AutoWriteConfig();
        commonWriteConfig.setThemeStyleRender(new AxolotlConfigurableTheme(AxolotlDefaultStyleConfig.class));
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING,true);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH,true);
        commonWriteConfig.setTitle("股票测试表");
        commonWriteConfig.setFontName("微软雅黑");
        commonWriteConfig.setBlankValue("-");
        commonWriteConfig.setOutputStream(fileOutputStream);
        List<Header> headers = new ArrayList<>();
        headers.add(new Header("代码","code"));
        headers.add(new Header("简称","intro"));
        headers.add(new Header("最新日期","localDateTimeStr"));
        headers.add(new Header("最新收盘价（元）",StockEntity::getClosingPrice));
        headers.add(new Header("涨跌幅（%）",StockEntity::getPriceLimit));
        headers.add(new Header("总市值（亿元）",StockEntity::getTotalValue));
        headers.add(new Header("流通市值（亿元）",StockEntity::getCirculationMarketValue));
        List<Header> subHeader1 = List.of(new Header("TTM"), new Header("15E"),new Header("16E"));
        headers.add(new Header("每股收益", subHeader1));
        headers.add(new Header("市盈率PE", subHeader1));
        headers.add(new Header("市净率PB（LF）"));
        headers.add(new Header("市销率PS（TTM）","pts"));
        ArrayList<StockEntity> data = new ArrayList<>();
        for (int i = 0; i < 50; i++) {
            StockEntity stockEntity = new StockEntity();
            stockEntity.setCode(RandomStringUtils.randomNumeric(8));
            StringBuilder sb = new StringBuilder();
            for (int i1 = 0; i1 < 10; i1++) {
                char c = RandomUtil.randomChinese();
                sb.append(c);
            }
            stockEntity.setIntro(sb.toString());
            stockEntity.setPts(Double.parseDouble(RandomStringUtils.randomNumeric(2)));
            data.add(stockEntity);
        }
        AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(commonWriteConfig);
        autoExcelWriter.write(headers,data);
        autoExcelWriter.close();

    }


    @Test
    public void testAuto3() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream("D:\\" + IdUtil.randomUUID() + ".xlsx");
        AutoWriteConfig commonWriteConfig = new AutoWriteConfig();
        commonWriteConfig.setThemeStyleRender(new AxolotlConfigurableTheme(new AxolotlDefaultStyleConfig()));
        //commonWriteConfig.setThemeStyleRender(ExcelWriteThemes.$DEFAULT);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_HIDDEN_BLANK_COLUMNS,false);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING,true);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH,true);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,true);
        commonWriteConfig.setTitle("——巡视总报告——");
      //  commonWriteConfig.setFontName("仿宋");
        commonWriteConfig.setBlankValue("-");
        commonWriteConfig.setOutputStream(fileOutputStream);
        List<Header> headers = new ArrayList<>();

        /*AxolotlCellStyle axolotlCellStyle = new AxolotlCellStyle();
        axolotlCellStyle.setForegroundColor(new AxolotlColor(155,123,147));
        Header header = new Header("杆塔及拉线", "one");
        header.setAxolotlCellStyle(axolotlCellStyle);
        headers.add(header);*/

        headers.add(new Header("导、地线","two")
                .axolotlCellStyle(
                        new AxolotlCellStyle()
                                .foregroundColor(new AxolotlColor(155,123,147))
                )
        );
        headers.add(new Header("绝缘子","three"));
        headers.add(new Header("金属附件及附属设备","four"));
        headers.add(new Header("基础及接地","five"));
        headers.add(new Header("通道","six"));

        ArrayList<JSONObject> data = new ArrayList<>();
       // AxolotlSelectBox<String> stringAxolotlSelectBox = new AxolotlSelectBox<>("阶段5", null);
        AxolotlSelectBox<LocalDate> stringAxolotlSelectBox = new AxolotlSelectBox<>(LocalDate.now(), List.of(LocalDate.now()));
        for (int i = 0; i < 50; i++) {
            JSONObject map = new JSONObject(true);
//            HashMap<String, String> map = new HashMap<>();
            map.put("two",stringAxolotlSelectBox);
            map.put("three","绝缘子XXX");
            map.put("four","金具XXX");
            map.put("five","基础XXX");
            map.put("six","通道XXX");
            map.put("one","杆塔XXX");
            data.add(map);
        }
        AxolotlAutoExcelWriter autoExcelWriter = Axolotls.getAutoExcelWriter(commonWriteConfig);
        SXSSFWorkbook workbook = autoExcelWriter.getWorkbook();
        autoExcelWriter.write(headers,data);
        autoExcelWriter.close();

    }


}

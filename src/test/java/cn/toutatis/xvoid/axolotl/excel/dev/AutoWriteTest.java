package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.hutool.core.lang.func.Func1;
import cn.hutool.core.util.IdUtil;
import cn.hutool.core.util.RandomUtil;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.AnnoEntity;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.StockEntity;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.support.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.ExcelWriteThemes;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.toolkit.clazz.LambdaToolkit;
import cn.toutatis.xvoid.toolkit.clazz.XFunc;
import com.alibaba.fastjson.JSONObject;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import java.io.*;
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
        SXSSFWorkbook workbook = autoExcelWriter.getWorkbook();
        autoExcelWriter.write(headers,data);
        autoExcelWriter.close();

    }

    @Test
    public void testAnno() throws FileNotFoundException {
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


}

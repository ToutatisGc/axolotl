package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.AnnoEntity;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.AxolotlColor;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.support.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.ExcelWriteThemes;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.CellStyle;
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
        commonWriteConfig.setStyleRender(ExcelWriteThemes.ADMINISTRATION_RED);
        commonWriteConfig.setWritePolicy(ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER,true);
        commonWriteConfig.setTitle("测试表");
        commonWriteConfig.setOutputStream(fileOutputStream);
        List<Header> headers = new ArrayList<>();
        headers.add(new Header("名\r\n称",List.of(new Header("姓名"),new Header("花名"))));
        headers.add(new Header("期限", List.of(new Header("年"), new Header("月"))));
        Header header1 = new Header("金额");
        Header header = new Header("账面数",
                List.of(new Header("经济",
                        List.of(new Header("数量"), new Header("金额"))), new Header("数量"), header1));
        headers.add(header);
        Header remark = new Header("备注");
        AxolotlCellStyle axolotlCellStyle = new AxolotlCellStyle();
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

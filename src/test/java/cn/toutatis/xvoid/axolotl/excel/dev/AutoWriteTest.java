package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.AnnoEntity;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import com.alibaba.fastjson.JSONObject;
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
        commonWriteConfig.setTitle("测试表");
        commonWriteConfig.setOutputStream(fileOutputStream);
        List<Header> headers = new ArrayList<>();
        headers.add(new Header("名称"));
        headers.add(new Header("期限", List.of(new Header("年"), new Header("月"))));
        headers.add(new Header("账面数",
                List.of(new Header("经济",
                        List.of(new Header("数量"), new Header("金额"))), new Header("非经济", List.of(new Header("数量"), new Header("金额"))))));
        headers.add(new Header("备注"));

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

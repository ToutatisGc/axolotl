package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.AnnoEntity;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import cn.toutatis.xvoid.axolotl.excel.writer.components.Header;
import com.alibaba.fastjson.JSONObject;
import org.junit.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class AutoWriteTest {


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
        int maxDepth = getMaxDepth(headers, 0);
        System.err.println(maxDepth);
        for (Header header : headers) {
            System.err.println(header.getTotalBottomLevelCount());
        }
    }

    public static int getMaxDepth(List<Header> headers, int depth) {
        int maxDepth = depth;
        for (Header header : headers) {
            if (header.getChilds() != null) {
                int subDepth = getMaxDepth(header.getChilds(), depth + 1);
                if (subDepth > maxDepth) {
                    maxDepth = subDepth;
                }
            }
        }
        return maxDepth;
    }

}

package cn.toutatis.xvoid.axolotl.excel.dev;

import cn.hutool.core.util.IdUtil;
import cn.toutatis.xvoid.axolotl.Axolotls;
import cn.toutatis.xvoid.axolotl.excel.entities.writer.AnnoEntity;
import cn.toutatis.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.AxolotlAutoExcelWriter;
import com.alibaba.fastjson.JSONObject;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
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

}
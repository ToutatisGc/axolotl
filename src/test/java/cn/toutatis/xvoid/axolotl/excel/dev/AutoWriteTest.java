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
        headers.add(new Header("账面数", List.of(new Header("经济",List.of(new Header("数量"),new Header("金额"))),new Header("非经济"))));
        int i = getMaxDepth(headers);
        setColumnRanges(headers);
        System.err.println(i);
    }

    public static int getMaxDepth(List<Header> headers) {
        if (headers == null || headers.isEmpty()) {
            return 0;
        }

        int maxDepth = 1;
        for (Header header : headers) {
            int depth = getMaxDepth(header, 1);
            maxDepth = Math.max(maxDepth, depth);
        }
        return maxDepth;
    }

    private static int getMaxDepth(Header header, int depth) {
        if (header.getChilds().isEmpty()) {
            return depth;
        }
        int maxChildDepth = depth;
        for (Header child : header.getChilds()) {
            int childDepth = getMaxDepth(child, depth + 1);
            maxChildDepth = Math.max(maxChildDepth, childDepth);
        }
        return maxChildDepth;
    }

    public static void setColumnRanges(List<Header> headers) {
        if (headers == null || headers.isEmpty()) {
            return;
        }

        for (Header header : headers) {
            setColumnRange(header);
        }
    }

    private static int setColumnRange(Header header) {
        if (header.getChilds().isEmpty()) {
            header.setColumnRange(1);
            return 1;
        }

        int childSize = header.getChilds().size();
        header.setColumnRange(childSize);
        return childSize;
    }

}

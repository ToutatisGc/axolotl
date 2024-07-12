package cn.xvoid.axolotl.excel.dev;

import cn.xvoid.axolotl.excel.entities.writer.DictTestEntity;
import cn.xvoid.axolotl.excel.entities.writer.DictTestFullEntity;
import cn.xvoid.axolotl.excel.writer.AutoWriteConfig;
import cn.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Map;

public class ConfigTest {

    @Test(expected = AxolotlWriteException.class)
    public void test1(){
        AutoWriteConfig autoWriteConfig = new AutoWriteConfig();
        ArrayList<DictTestEntity> dict = new ArrayList<>();
        dict.add(new DictTestEntity("TES_001", "状态正常"));
        dict.add(new DictTestEntity("TES_002", "状态异常"));
        dict.add(new DictTestEntity("TES_003", "状态失效"));
        autoWriteConfig.setDict(0,"status", dict);
    }

    @Test
    public void test2(){
        AutoWriteConfig autoWriteConfig = new AutoWriteConfig();
        ArrayList<DictTestFullEntity> dict = new ArrayList<>();
        dict.add(new DictTestFullEntity("TES_001", "状态正常"));
        dict.add(new DictTestFullEntity("TES_002", "状态异常"));
        dict.add(new DictTestFullEntity("TES_003", "状态失效"));
        autoWriteConfig.setDict(0,"status", dict);
        autoWriteConfig.setDict(0,"payStatus",Map.of(
                "PAY_001","拉起支付",
                "PAY_002","支付成功",
                "PAY_003","取消支付"));
        System.err.println(autoWriteConfig.getDictionaryMapping());

        var w = autoWriteConfig.getDict(0, "status1");
        System.err.println(w);

    }

}

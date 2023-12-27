package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.entities.IndexTest;
import cn.toutatis.xvoid.axolotl.support.WorkBookReaderConfig;
import org.junit.Test;

public class ReadConfigTest {

    @Test
    public void testWorkBookReaderConfig() {
        WorkBookReaderConfig<IndexTest> config = new WorkBookReaderConfig<>();
        config.setCastClass(IndexTest.class);
        for (EntityCellMappingInfo entityCellMappingInfo : config.getIndexMappingInfos()) {
            System.err.println(entityCellMappingInfo);
        }

    }
}

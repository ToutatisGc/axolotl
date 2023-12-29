package cn.toutatis.xvoid.axolotl.dev;

import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.entities.IndexTest;
import cn.toutatis.xvoid.axolotl.excel.ReaderConfig;
import org.junit.Test;

public class ReadConfigTest {

    @Test
    public void testWorkBookReaderConfig() {
        ReaderConfig<IndexTest> config = new ReaderConfig<>();
        config.setCastClass(IndexTest.class);
        for (EntityCellMappingInfo entityCellMappingInfo : config.getIndexMappingInfos()) {
            System.err.println(entityCellMappingInfo);
        }

    }
}

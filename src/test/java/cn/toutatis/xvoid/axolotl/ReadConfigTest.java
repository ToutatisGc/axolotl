package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.entities.IndexPropertyEntity;
import cn.toutatis.xvoid.axolotl.support.WorkBookReaderConfig;
import org.junit.Test;

public class ReadConfigTest {

    @Test
    public void testWorkBookReaderConfig() {
        WorkBookReaderConfig<IndexPropertyEntity> config = new WorkBookReaderConfig<>();
        config.setCastClass(IndexPropertyEntity.class);
    }
}

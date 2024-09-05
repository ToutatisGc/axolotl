package cn.xvoid.axolotl.excel.dev;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.support.docker.MapDocker;
import cn.xvoid.common.exception.base.VoidRuntimeException;
import cn.xvoid.toolkit.clazz.ClassToolkit;
import org.junit.Test;

import java.io.IOException;
import java.util.List;

public class FeatureTest {

    @Test
    public void testPackage(){
        String packageName = MapDocker.class.getPackageName();
        System.err.println(packageName);
        try {
            List<Class<?>> classesForPackage = ClassToolkit.getClassesForPackage(packageName+".impl");
            System.err.println(classesForPackage);
        } catch (IOException | ClassNotFoundException e) {
            throw new VoidRuntimeException(e.getMessage());
        }
    }

    @Test
    public void testConfig() {
        ReaderConfig<Object> objectReaderConfig = new ReaderConfig<>();
        System.err.println(objectReaderConfig);
    }

}

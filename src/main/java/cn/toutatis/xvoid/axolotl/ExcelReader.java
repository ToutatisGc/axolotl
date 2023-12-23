package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.support.DetectResult;
import cn.toutatis.xvoid.axolotl.support.TikaShell;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import org.slf4j.Logger;

import java.io.File;

public class ExcelReader {

    private static final Logger LOGGER  = LoggerToolkit.getLogger(ExcelReader.class);

    public ExcelReader(File excelFile) {
        DetectResult detect = TikaShell.detect(excelFile, TikaShell.OOXML_EXCEL);
//        if (detect.isDetect())
    }
}

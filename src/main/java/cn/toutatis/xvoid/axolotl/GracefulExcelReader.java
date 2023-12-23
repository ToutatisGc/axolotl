package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.support.DetectResult;
import cn.toutatis.xvoid.axolotl.support.TikaShell;
import cn.toutatis.xvoid.axolotl.support.WorkBookMetaInfo;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import lombok.Getter;
import org.slf4j.Logger;

import java.io.File;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class GracefulExcelReader {

    private final Logger LOGGER  = LoggerToolkit.getLogger(GracefulExcelReader.class);

    @Getter
    private WorkBookMetaInfo workBookMetaInfo;

    public GracefulExcelReader(File excelFile) {
        DetectResult detectResult = TikaShell.detectThrowException(excelFile, TikaShell.OOXML_EXCEL);
//        if (detectResult.isWantedMimeType()){
//
//        }else if (!detectResult.isDetect())
//        if (detectResult.isDetect()){
//
//        }else {
//            TikaShell.detect(excelFile, TikaShell.MS_EXCEL);
//        }
//        this.workBookMetaInfo = new WorkBookMetaInfo(excelFile);

    }
}

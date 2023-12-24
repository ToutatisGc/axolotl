package cn.toutatis.xvoid.axolotl.support;

import java.io.File;

public class WorkBookMetaInfo extends AbstractMetaInfo{
    public WorkBookMetaInfo(File file, DetectResult detectResult) {
        this.setFile(file);
        this.setMimeType(detectResult.getCatchMimeType());
    }


    //    private List

}

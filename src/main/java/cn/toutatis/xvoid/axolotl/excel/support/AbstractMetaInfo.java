package cn.toutatis.xvoid.axolotl.excel.support;

import lombok.Data;
import org.apache.tika.mime.MimeType;

import java.io.File;

/**
 * 文件元信息抽象类
 */
@Data
public abstract class AbstractMetaInfo {

    private File file;

    private String originFileName;

    private MimeType mimeType;

    public void setFile(File file) {
        this.file = file;
        this.originFileName = file.getName();
    }
}

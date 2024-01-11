package cn.toutatis.xvoid.axolotl.excel.reader.support;

import lombok.Data;
import org.apache.tika.mime.MimeType;

import java.io.File;

/**
 * 文件元信息抽象类
 */
@Data
public abstract class AbstractContext {

    private File file;

    private String originFileName;

    private MimeType mimeType;

    public void setFile(File file) {
        if (file == null){
            throw new IllegalArgumentException("文件不得为空");
        }
        this.file = file;
        this.originFileName = file.getName();
    }
}

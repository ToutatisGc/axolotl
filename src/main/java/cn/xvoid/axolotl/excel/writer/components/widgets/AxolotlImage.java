package cn.xvoid.axolotl.excel.writer.components.widgets;

import cn.xvoid.axolotl.common.CommonMimeType;
import cn.xvoid.axolotl.toolkit.tika.DetectResult;
import cn.xvoid.axolotl.toolkit.tika.TikaShell;
import cn.xvoid.toolkit.constant.Regex;
import cn.xvoid.toolkit.file.FileToolkit;
import cn.xvoid.toolkit.validator.Validator;
import lombok.Data;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tika.mime.MimeType;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Base64;
import java.util.regex.Pattern;

@Data
public class AxolotlImage {

    private static final Pattern IMAGE_PATTERN = Pattern.compile(Regex.IMAGE_SUFFIX_REGEX);

    private byte[] data;

    private ClientAnchor.AnchorType anchorType = ClientAnchor.AnchorType.MOVE_AND_RESIZE;

    private int imageFormat = XSSFWorkbook.PICTURE_TYPE_JPEG;

    public AxolotlImage(byte[] data,int imageFormat) {
        if (data == null || data.length == 0){throw new IllegalArgumentException("数据不得为空");}
        this.data = data;
        this.imageFormat = imageFormat;
    }

    public AxolotlImage(File file) {
        DetectResult imageFormatDetectResult = TikaShell.detect(file, CommonMimeType.JPEG);
        if (!imageFormatDetectResult.isDetect()){
            imageFormatDetectResult = TikaShell.detect(file, CommonMimeType.PNG,true);
        }else{
            imageFormat = XSSFWorkbook.PICTURE_TYPE_JPEG;
        }
        if (imageFormatDetectResult.isDetect()){
            this.data = FileToolkit.getFileBytes(file);
            this.imageFormat = XSSFWorkbook.PICTURE_TYPE_PNG;
        }else {
            throw new IllegalArgumentException("文件格式错误,支持[JPEG,PNG]");
        }
    }

    public AxolotlImage(InputStream inputStream,int imageFormat) throws IOException {
        if (inputStream == null){throw new IllegalArgumentException("输入流为空");}
        this.data = IOUtils.toByteArray(inputStream);
        this.imageFormat = imageFormat;
    }

    public AxolotlImage(String base64,int imageFormat) throws IOException {
        if (Validator.strIsBlank(base64)){throw new IllegalArgumentException("base64串为空");}
        this.data = Base64.getDecoder().decode(base64.getBytes());
        this.imageFormat = imageFormat;
    }

    public void checkImage(){
         if (!(this.data != null && this.data.length > 0)){throw new IllegalArgumentException("图片数据为空");}
         if (imageFormat != XSSFWorkbook.PICTURE_TYPE_JPEG && imageFormat != XSSFWorkbook.PICTURE_TYPE_PNG){
             throw new IllegalArgumentException("图片格式错误,仅支持[JPEG,PNG]");
         }
    }

}

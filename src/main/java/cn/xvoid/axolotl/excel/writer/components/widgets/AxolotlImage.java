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

/**
 * AxolotlImage 类用于表示 Excel 文档中的图像组件，并支持图像插入功能。
 * 该功能目前并不支持图片的offset偏移设置，目前无需求。
 * @see cn.xvoid.axolotl.excel.writer.AxolotlTemplateExcelWriter#writeImage(int, AxolotlImage) 
 * @author Toutatis_Gc
 */
@Data
public class AxolotlImage {

    /**
     * 图像后缀正则表达式模式。
     */
    private static final Pattern IMAGE_PATTERN = Pattern.compile(Regex.IMAGE_SUFFIX_REGEX);

    /**
     * 图像数据。
     */
    private byte[] data;

    /**
     * 图像锚点类型，默认为移动并调整大小。
     */
    private ClientAnchor.AnchorType anchorType = ClientAnchor.AnchorType.MOVE_AND_RESIZE;

    /**
     * 图像格式。
     */
    private int imageFormat;

    /**
     * 图像在 Excel 表格中的起始列号。
     */
    private int startColumn;

    /**
     * 图像在 Excel 表格中的起始行号。
     */
    private int startRow;

    /**
     * 图像在 Excel 表格中的结束列号。
     */
    private int endColumn;

    /**
     * 图像在 Excel 表格中的结束行号。
     */
    private int endRow;

    /**
     * 构造一个带有字节数组数据的图像对象。
     *
     * @param data 图像数据作为字节数组
     * @param imageFormat 图像格式类型
     * @throws IllegalArgumentException 如果数据为空或长度为零
     */
    public AxolotlImage(byte[] data, int imageFormat) {
        if (data == null || data.length == 0) { throw new IllegalArgumentException("数据不得为空"); }
        this.data = data;
        this.imageFormat = imageFormat;
    }

    /**
     * 构造一个带有 File 对象的图像对象，自动检测图像格式。
     *
     * @param file 图像文件
     * @throws IllegalArgumentException 如果文件格式不被支持
     */
    public AxolotlImage(File file) {
        DetectResult imageFormatDetectResult = TikaShell.detect(file, CommonMimeType.JPEG);
        if (!imageFormatDetectResult.isDetect()) {
            imageFormatDetectResult = TikaShell.detect(file, CommonMimeType.PNG, true);
        } else {
            imageFormat = XSSFWorkbook.PICTURE_TYPE_JPEG;
        }
        if (imageFormatDetectResult.isDetect()) {
            this.data = FileToolkit.getFileBytes(file);
            this.imageFormat = XSSFWorkbook.PICTURE_TYPE_PNG;
        } else {
            throw new IllegalArgumentException("文件格式错误,支持[JPEG,PNG]");
        }
    }

    /**
     * 构造一个带有 InputStream 的图像对象。
     *
     * @param inputStream 包含图像数据的输入流
     * @param imageFormat 图像格式类型
     * @throws IOException 如果发生 I/O 错误
     */
    public AxolotlImage(InputStream inputStream, int imageFormat) throws IOException {
        if (inputStream == null) { throw new IllegalArgumentException("输入流为空"); }
        this.data = IOUtils.toByteArray(inputStream);
        this.imageFormat = imageFormat;
    }

    /**
     * 构造一个带有 Base64 编码字符串的图像对象。
     *
     * @param base64 Base64 编码的图像数据字符串
     * @param imageFormat 图像格式类型
     * @throws IOException 如果发生 I/O 错误
     */
    public AxolotlImage(String base64, int imageFormat) throws IOException {
        if (Validator.strIsBlank(base64)) { throw new IllegalArgumentException("base64串为空"); }
        this.data = Base64.getDecoder().decode(base64.getBytes());
        this.imageFormat = imageFormat;
    }

    /**
     * 设置起始和结束位置
     *
     * @param startColumn 起点的列坐标
     * @param startRow 起点的行坐标
     * @param endColumn 终点的列坐标
     * @param endRow 终点的行坐标
     */
    public void setPosition(int startColumn, int startRow, int endColumn, int endRow) {
        this.startColumn = startColumn;
        this.startRow = startRow;
        this.endColumn = endColumn;
        this.endRow = endRow;
    }

    /**
     * 验证图像数据和格式，确保它们符合指定要求。
     *
     * @throws IllegalArgumentException 如果图像数据为空或格式不被支持
     */
    public void checkImage() {
        if (!(this.data != null && this.data.length > 0)) { throw new IllegalArgumentException("图片数据为空"); }
        if (imageFormat != XSSFWorkbook.PICTURE_TYPE_JPEG && imageFormat != XSSFWorkbook.PICTURE_TYPE_PNG) {
            throw new IllegalArgumentException("图片格式错误,仅支持[JPEG,PNG]");
        }
    }

}

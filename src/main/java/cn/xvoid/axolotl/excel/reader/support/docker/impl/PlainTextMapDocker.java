package cn.xvoid.axolotl.excel.reader.support.docker.impl;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.xvoid.axolotl.excel.reader.support.docker.AbstractMapDocker;
import cn.xvoid.common.standard.StringPool;
import cn.xvoid.toolkit.constant.Regex;
import cn.xvoid.toolkit.number.NumberToolkit;
import org.apache.poi.ss.usermodel.CellType;

import java.text.DecimalFormat;

/**
 * 纯文本映射Docker实现类
 * 用于处理Excel单元格中的纯文本数据
 * @author Toutatis_Gc
 */
public class PlainTextMapDocker extends AbstractMapDocker<String> {

    /**
     * 定义该Docker的后缀名
     */
    public static final String SUFFIX_NAME = "PLAIN_TEXT";

    /**
     * 用于格式化数字的DecimalFormat实例
     */
    private final DecimalFormat decimalTextFormat = new DecimalFormat(StringPool.HASH);

    /**
     * 获取该Docker的后缀名
     *
     * @return Docker的后缀名
     */
    @Override
    public String getSuffix() {
        return SUFFIX_NAME;
    }

    /**
     * 转换Excel单元格值为纯文本
     *
     * @param index 单元格索引
     * @param cellGetInfo 单元格获取信息
     * @param readerConfig 读取配置
     * @return 转换后的纯文本字符串
     */
    @Override
    public String convert(int index, CellGetInfo cellGetInfo, ReaderConfig<?> readerConfig) {
        Object cellValue = cellGetInfo.getCellValue();
        if (cellValue == null) {
            return null;
        }
        String formatValue;
        if (cellGetInfo.getCellType() == CellType.NUMERIC) {
            // 使用预定义的格式化规则处理数值型单元格
            if (NumberToolkit.isInteger((Double) cellValue)){
                formatValue = decimalTextFormat.format(cellValue);
            }else{
                formatValue = cellValue.toString();
            }
        } else {
            // 直接转换非数值型单元格为字符串
            formatValue = cellValue.toString();
        }
        if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.TRIM_CELL_VALUE)) {
            // 根据读取策略，可能需要去除单元格值的前后空格
            formatValue = Regex.convertSingleLine(formatValue).replace(" ", "");
        }
        return formatValue;
    }
}

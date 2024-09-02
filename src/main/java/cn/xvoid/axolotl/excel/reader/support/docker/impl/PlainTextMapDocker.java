package cn.xvoid.axolotl.excel.reader.support.docker.impl;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.xvoid.axolotl.excel.reader.support.docker.AbstractMapDocker;
import cn.xvoid.common.standard.StringPool;
import cn.xvoid.toolkit.constant.Regex;
import org.apache.poi.ss.usermodel.CellType;

import java.text.DecimalFormat;

public class PlainTextMapDocker extends AbstractMapDocker<String> {

    public static final String SUFFIX_NAME = "PLAIN_TEXT";

    private final DecimalFormat decimalTextFormat = new DecimalFormat(StringPool.HASH);

    @Override
    public String getSuffix() {return SUFFIX_NAME;}

    @Override
    public String convert(int index, CellGetInfo cellGetInfo, ReaderConfig<?> readerConfig) {
        Object cellValue = cellGetInfo.getCellValue();
        if (cellValue == null){
            return null;
        }
        String formatValue;
        if (cellGetInfo.getCellType() == CellType.NUMERIC){
            formatValue = decimalTextFormat.format(cellValue);
        }else {
            formatValue = cellValue.toString();
        }
        if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.TRIM_CELL_VALUE)){
            formatValue = Regex.convertSingleLine(formatValue).replace(" ", "");
        }
        return formatValue;
    }
}

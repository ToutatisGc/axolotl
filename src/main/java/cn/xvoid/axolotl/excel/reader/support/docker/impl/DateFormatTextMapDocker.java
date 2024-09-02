package cn.xvoid.axolotl.excel.reader.support.docker.impl;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.xvoid.axolotl.excel.reader.support.docker.AbstractMapDocker;
import cn.xvoid.toolkit.constant.Time;
import cn.xvoid.toolkit.log.LoggerToolkit;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.slf4j.Logger;

public class DateFormatTextMapDocker extends AbstractMapDocker<String> {

    public static final String SUFFIX_NAME = "DATE_FMT";

    private static final Logger LOGGER = LoggerToolkit.getLogger(DateFormatTextMapDocker.class);

    @Override
    public String getSuffix() {
        return SUFFIX_NAME;
    }

    @Override
    public String convert(int index, CellGetInfo cellGetInfo, ReaderConfig<?> readerConfig) {
        Object cellValue = cellGetInfo.getCellValue();
        if (cellValue == null){
            return null;
        }
        CellType cellType = cellGetInfo.getCellType();
        if (cellType == CellType.NUMERIC){
            Cell cell = cellGetInfo.get_cell();
            if (DateUtil.isCellDateFormatted(cell)){
                return Time.regexTime(cell.getDateCellValue());
            }else{
                if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.USE_MAP_DEBUG) && LOGGER.isTraceEnabled()){
                    LOGGER.trace("日期格式化失败，仅为数字格式可转换["+cellValue+"]");
                }
                return null;
            }
        }
        return null;
    }
}

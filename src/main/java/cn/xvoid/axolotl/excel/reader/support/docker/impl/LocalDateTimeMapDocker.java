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

import java.time.LocalDateTime;
import java.util.Date;

/**
 * 日期格式文本映射停靠类，继承自抽象映射停靠类
 * 专门处理日期格式转换为文本的操作
 * @author Toutatis_Gc
 */
public class LocalDateTimeMapDocker extends AbstractMapDocker<LocalDateTime> {

    /**
     * 无参构造方法，默认为不显示空值
     */
    public LocalDateTimeMapDocker() {
        this.setNullDisplay(false);
    }

    // 定义后缀名为日期格式的常量
    public static final String SUFFIX_NAME = "LOCAL_DATE_TIME";

    // 初始化日志记录器，用于记录类的操作日志
    private static final Logger LOGGER = LoggerToolkit.getLogger(LocalDateTimeMapDocker.class);

    /**
     * 获取后缀名
     *
     * @return 返回后缀名SUFFIX_NAME
     */
    @Override
    public String getSuffix() {
        return SUFFIX_NAME;
    }

    /**
     * 转换单元格值为日期格式的文本
     *
     * @param index 单元格索引，未使用
     * @param cellGetInfo 单元格获取信息对象，包含单元格的值和类型
     * @param readerConfig 读取器配置对象，决定是否使用调试策略
     * @return 返回转换后的日期格式文本，如果无法转换则返回null
     */
    @Override
    public LocalDateTime convert(int index, CellGetInfo cellGetInfo, ReaderConfig<?> readerConfig) {
        // 获取单元格的值
        Object cellValue = cellGetInfo.getCellValue();
        // 如果单元格值为空，直接返回null
        if (cellValue == null){
            return null;
        }
        // 获取单元格的类型
        CellType cellType = cellGetInfo.getCellType();
        // 如果单元格类型为数值
        if (cellType == CellType.NUMERIC){
            // 获取单元格对象
            Cell cell = cellGetInfo.get_cell();
            // 如果单元格是以日期格式显示的数值
            if (DateUtil.isCellDateFormatted(cell)){
                // 返回格式化后的日期字符串
                return DateUtil.getLocalDateTime((Double) cellGetInfo.getCellValue());
            }else{
                // 如果启用了调试策略并且日志记录器允许trace级别日志
                if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.USE_MAP_DEBUG) && LOGGER.isTraceEnabled()){
                    // 记录转换失败的日志，说明单元格值仅为数字，无法转换为日期格式
                    LOGGER.trace("日期格式化失败，仅为数字并且为日期格式可转换["+cellValue+"]");
                }
                // 如果单元格值不是日期格式的数值，返回null
                return null;
            }
        }
        // 如果单元格类型不是数值，返回null
        return null;
    }
}

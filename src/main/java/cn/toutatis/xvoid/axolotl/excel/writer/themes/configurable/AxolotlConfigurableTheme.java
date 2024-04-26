package cn.toutatis.xvoid.axolotl.excel.writer.themes.configurable;

import cn.toutatis.xvoid.axolotl.excel.writer.components.configuration.AxolotlCellStyle;
import cn.toutatis.xvoid.axolotl.excel.writer.components.widgets.Header;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.*;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.AxolotlWriteResult;
import cn.toutatis.xvoid.axolotl.excel.writer.support.inverters.DataInverter;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Data;
import lombok.SneakyThrows;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;

import java.io.Serializable;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.*;

import static cn.toutatis.xvoid.axolotl.excel.writer.style.StyleHelper.START_POSITION;
import static cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.*;
import static cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper.debug;

/**
 * 可配置主题
 * @author 张智凯
 * @version 1.0
 * @data 2024/3/28 9:21
 */
public class AxolotlConfigurableTheme extends AbstractStyleRender implements ExcelStyleRender {

    /**
     * 标题默认行高
     */
    public static final Short TITLE_ROW_HEIGHT = StyleHelper.STANDARD_TITLE_ROW_HEIGHT;

    /**
     * 表头默认行高
     */
    public static final Short HEADER_ROW_HEIGHT = StyleHelper.STANDARD_ROW_HEIGHT;

    /**
     * 内容默认行高
     */
    public static final Short DATA_ROW_HEIGHT = StyleHelper.STANDARD_ROW_HEIGHT;

    /**
     * 默认列宽
     */
    public static final Short DEFAULT_COLUMN_WIDTH = (short) 12;

    /**
     * 日志
     */
    private static final Logger LOGGER = LoggerToolkit.getLogger(AxolotlConfigurableTheme.class);

    /**
     * 全局单元格属性
     */
    private CellPropertyHolder globalCellPropHolder = CellPropertyHolder.buildDefault();

    /**
     * 样式创建缓存
     */
    private final Map<CellPropertyHolder,CellStyle> cellStyleCache = new HashMap<>();

    /**
     * 表头单元格属性
     */
    private CellPropertyHolder headerCellPropHolder;

    /**
     * 标题单元格属性
     */
    private CellPropertyHolder titleCellPropHolder;

    /**
     * 程序写入单元格属性
     */
    private Map<ExcelWritePolicy,CellPropertyHolder> commonCellPropHolder;

    /**
     * 样式配置类
     */
    private final ConfigurableStyleConfig configurableStyleConfig;

    public AxolotlConfigurableTheme() {
        super(LOGGER);
        this.configurableStyleConfig = new ConfigurableStyleConfig() {};
    }

    public AxolotlConfigurableTheme(ConfigurableStyleConfig configurableStyleConfig) {
        super(LOGGER);
        if(configurableStyleConfig == null){
            throw new AxolotlWriteException("无法加载样式配置，请传入配置类");
        }
        this.configurableStyleConfig = configurableStyleConfig;
    }

    public AxolotlConfigurableTheme(Class<? extends ConfigurableStyleConfig> configurClass) {
        super(LOGGER);
        if(configurClass == null){
            throw new AxolotlWriteException("无法加载样式配置，请传入配置类");
        }
        try {
            this.configurableStyleConfig = configurClass.getDeclaredConstructor().newInstance();
        } catch (InstantiationException | IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public AxolotlWriteResult init(SXSSFSheet sheet) {
        AxolotlWriteResult axolotlWriteResult;
        if(isFirstBatch()){
            CellConfigProperty globalConfig = new CellConfigProperty();
            configurableStyleConfig.globalStyleConfig(globalConfig);
            globalCellPropHolder = ConfigurableStyleConfig.cloneStyleProperties(globalConfig,globalCellPropHolder);
            String fontName = writeConfig.getFontName();
            if (fontName != null){
                debug(LOGGER, "使用自定义字体：%s",fontName);
                globalCellPropHolder.setFontName(fontName);
            }
            debug(LOGGER,"全局样式读取完毕");

            //程序写入单元格样式配置
            commonCellPropHolder = ConfigurableStyleConfig.loadCommonConfigFromDefault(globalCellPropHolder,configurableStyleConfig);
            debug(LOGGER,"程序写入单元格样式读取完毕");

            axolotlWriteResult = new AxolotlWriteResult(true,"初始化成功");
            String sheetName = writeConfig.getSheetName();
            if(Validator.strNotBlank(sheetName)){
                int sheetIndex = context.getSwitchSheetIndex();
                info(LOGGER,"设置工作表索引[%s]表名为:[%s]",sheetIndex,sheetName);
                context.getWorkbook().setSheetName(sheetIndex,sheetName);
            }else {
                debug(LOGGER,"未设置工作表名称");
            }
            boolean fillWhite = writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_FILL_DEFAULT_CELL_WHITE);
            if (fillWhite){
                this.fillWhiteCell(sheet);
            }
        }else {
            axolotlWriteResult = new AxolotlWriteResult(true,"已初始化");
        }
        return axolotlWriteResult;
    }

    @Override
    public AxolotlWriteResult renderHeader(SXSSFSheet sheet) {
        //读取表头配置
        CellConfigProperty headerConfig = new CellConfigProperty();
        configurableStyleConfig.headerStyleConfig(headerConfig);
        headerCellPropHolder = ConfigurableStyleConfig.cloneStyleProperties(headerConfig,globalCellPropHolder);
        if(headerCellPropHolder.getRowHeight() == null){
            headerCellPropHolder.setRowHeight(HEADER_ROW_HEIGHT);
        }
        debug(LOGGER,"表头样式读取完毕");

        //读取标题配置
        CellConfigProperty titleConfig = new CellConfigProperty();
        configurableStyleConfig.titleStyleConfig(titleConfig);
        titleCellPropHolder = ConfigurableStyleConfig.cloneStyleProperties(titleConfig,globalCellPropHolder);
        if(titleCellPropHolder.getRowHeight() == null){
            titleCellPropHolder.setRowHeight(TITLE_ROW_HEIGHT);
        }
        debug(LOGGER,"标题样式读取完毕");

        // 1.创建标题行
        AxolotlWriteResult writeTitle = createTitleRow(sheet);

        // 2.渲染表头
        AxolotlWriteResult headerWriteResult = this.defaultRenderHeaders(sheet);

        // 3.合并标题  长度为表头的长度
        if (writeTitle.isWrite()){
            this.mergeTitleRegion(sheet,context.getAlreadyWrittenColumns().get(context.getSwitchSheetIndex()),createCellStyle(titleCellPropHolder));
        }

        // 4.创建冻结窗格
        sheet.createFreezePane(START_POSITION, context.getAlreadyWriteRow().get(context.getSwitchSheetIndex())+1);

        return headerWriteResult;
    }

    private boolean alreadyNotice = false;
    @Override
    public AxolotlWriteResult renderData(SXSSFSheet sheet, List<?> data) {
        //创建系统列单元格样式 用于序号与空值填充
        CellPropertyHolder cellProperty = commonCellPropHolder.get(AUTO_INSERT_SERIAL_NUMBER);
     //   XSSFCellStyle commonCellStyle = (XSSFCellStyle) createCellStyle(commonCellPropHolder);
        for (Object datum : data) {
            // 获取对象属性
            HashMap<String, Object> dataMap = this.getDataMap(datum);
            // 初始化内容
            HashMap<Integer, Integer> writtenColumnMap = new HashMap<>();
            int switchSheetIndex = getContext().getSwitchSheetIndex();
            Map<Integer, Integer> alreadyWriteRowMap = context.getAlreadyWriteRow();
            int alreadyWriteRow = alreadyWriteRowMap.getOrDefault(switchSheetIndex,-1);
            alreadyWriteRowMap.put(switchSheetIndex,++alreadyWriteRow);
            SXSSFRow dataRow = sheet.createRow(alreadyWriteRow);
            int writtenColumn = START_POSITION;
            int serialNumber = context.getAndIncrementSerialNumber() - context.getHeaderRowCount().get(switchSheetIndex);
            // 写入数据
            Map<String, Integer> columnMapping = context.getHeaderColumnIndexMapping().row(context.getSwitchSheetIndex());
            unmappedColumnCount =  new HashMap<>();
            columnMapping.forEach((key, value) -> unmappedColumnCount.put(value, 1));
            boolean columnMappingEmpty = columnMapping.isEmpty();
            // 写入序号
            //序号列索引
            int serialNumberColumnNumber = -1;
            boolean autoInsertSerialNumber = writeConfig.getWritePolicyAsBoolean(AUTO_INSERT_SERIAL_NUMBER);
            if (autoInsertSerialNumber){
                SXSSFCell cell = dataRow.createCell(writtenColumn);
                cell.setCellValue(serialNumber);
                if(cellProperty != null){
                    //设置行高 执行顺序为1 优先级较低
                    dataRow.setHeight(cellProperty.getRowHeight());
                    cell.setCellStyle(createCellStyle(cellProperty));
                }else{
                    serialNumberColumnNumber = writtenColumn;
                }
                writtenColumnMap.put(columnMappingEmpty?writtenColumn++:0,1);
            }
            for (Map.Entry<String, Object> dataEntry : dataMap.entrySet()) {
                String fieldName = dataEntry.getKey();
                SXSSFCell cell;
                if (columnMappingEmpty){
                    cell = dataRow.createCell(writtenColumn);
                }else{
                    if (columnMapping.containsKey(fieldName)){
                        cell = (SXSSFCell) ExcelToolkit.createOrCatchCell(sheet,alreadyWriteRow,columnMapping.get(fieldName),null);
                    }else {
                        if (!alreadyNotice){
                            warn(LOGGER,"未映射字段[%s]请在表头Header中映射字段!",fieldName);
                            alreadyNotice = true;
                        }
                        continue;
                    }
                }
                Object value = dataEntry.getValue();
                int columnNumber = columnMappingEmpty ? writtenColumn : columnMapping.get(fieldName);

                //读取内容样式配置
                CellConfigProperty dataConfig = new CellConfigProperty();
                configurableStyleConfig.dataStyleConfig(dataConfig,new FieldInfo(datum, fieldName, value,switchSheetIndex, columnNumber, alreadyWriteRow));
                CellPropertyHolder dataCellPropHolder = ConfigurableStyleConfig.cloneStyleProperties(dataConfig, globalCellPropHolder);
                if(dataCellPropHolder.getRowHeight() == null){
                    dataCellPropHolder.setRowHeight(DATA_ROW_HEIGHT);
                }
                List<Header> headers = context.getHeaders().get(switchSheetIndex);
                if(dataCellPropHolder.getColumnWidth() == null && (headers == null || headers.isEmpty())){
                    dataCellPropHolder.setColumnWidth(DEFAULT_COLUMN_WIDTH);
                }

                //设置行高
                dataRow.setHeight(dataCellPropHolder.getRowHeight());
                //设置列宽 内容列宽若没设置默认值，继承表头
                if(dataCellPropHolder.getColumnWidth() != null){
                    sheet.setColumnWidth(columnNumber,dataCellPropHolder.getColumnWidth());
                }

                // 对单元格设置样式
                cell.setCellStyle(createCellStyle(dataCellPropHolder));
                FieldInfo fieldInfo = new FieldInfo(datum, fieldName, value,switchSheetIndex, columnNumber ,alreadyWriteRow);
                // 渲染数据到单元格
//                this.renderColumn(sheet,fieldInfo,cell,unmappedColumnCount);
                this.renderColumn(fieldInfo,cell);
                unmappedColumnCount.remove(fieldInfo.getColumnIndex());
                if (columnMappingEmpty){
                    writtenColumnMap.put(writtenColumn++,1);
                }else{
                    writtenColumnMap.put(columnNumber,1);
                }
            }
            if(serialNumberColumnNumber != -1){
                //没有指定 程序写入单元格样式
                if(!dataMap.isEmpty()){
                    //取编号后第一个单元格的样式
                    CellStyle cellStyle = dataRow.getCell(serialNumberColumnNumber+1).getCellStyle();
                    dataRow.getCell(serialNumberColumnNumber).setCellStyle(cellStyle);
                }else{
                    //编号后第一个单元格没有数据 取全局样式
                    dataRow.getCell(serialNumberColumnNumber).setCellStyle(createCellStyle(globalCellPropHolder));
                }
            }
            // 将未使用的的单元格赋予空值
            for (int alreadyColumnIdx = 0; alreadyColumnIdx < context.getAlreadyWrittenColumns().get(switchSheetIndex); alreadyColumnIdx++) {
                if (autoInsertSerialNumber && alreadyColumnIdx == 0){
                    continue;
                }
                SXSSFCell cell = null;
                if (columnMappingEmpty){
                    if (!writtenColumnMap.containsKey(alreadyColumnIdx)){
                        cell = dataRow.createCell(alreadyColumnIdx);
                    }
                }else{
                    if (!columnMapping.containsValue(alreadyColumnIdx)){
                        cell = dataRow.createCell(alreadyColumnIdx);
                    }
                    if (unmappedColumnCount.containsKey(alreadyColumnIdx)){
                        cell = dataRow.createCell(alreadyColumnIdx);
                    }
                }
                if (cell != null){
                    cell.setCellValue(writeConfig.getBlankValue());
                    if(cellProperty != null){
                        cell.setCellStyle(createCellStyle(cellProperty));
                    }else{
                        cell.setCellStyle(createCellStyle(globalCellPropHolder));
                    }
                    //空值填充不设置行高列宽
                }
            }
            debug(LOGGER,"第["+alreadyWriteRow+"]行内容渲染结束");
        }


        return new AxolotlWriteResult(true, "渲染数据完成");
    }

    @Override
    public AxolotlWriteResult finish(SXSSFSheet sheet) {
        debug(LOGGER,"结束渲染工作表[%s]",sheet.getSheetName());
        int sheetIndex = context.getWorkbook().getSheetIndex(sheet);
        CellPropertyHolder cellProperty = commonCellPropHolder.get(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING);
        int alreadyWrittenColumns = context.getAlreadyWrittenColumns().get(sheetIndex);
        // 创建结尾合计行
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING)){
            Map<Integer, BigDecimal> endingTotalMapping = context.getEndingTotalMapping().row(sheetIndex);
            debug(LOGGER,"开始创建结尾合计行,合计数据为:%s",endingTotalMapping);
            SXSSFRow row = sheet.createRow(sheet.getLastRowNum() + 1);
            if(cellProperty == null){
                //未配置 取上一行行高
                row.setHeight(sheet.getRow(sheet.getLastRowNum() - 1).getHeight());
            }else{
                row.setHeight(cellProperty.getRowHeight());
            }
            DataInverter<?> dataInverter = writeConfig.getDataInverter();
            Set<Integer> calculateColumnIndexes = writeConfig.getCalculateColumnIndexes(sheetIndex);
            for (int i = 0; i < alreadyWrittenColumns; i++) {
                SXSSFCell cell = row.createCell(i);
                String cellValue = writeConfig.getBlankValue();
                if (i == 0 && writeConfig.getWritePolicyAsBoolean(AUTO_INSERT_SERIAL_NUMBER)){
                    cellValue = "合计";
                   // sheet.setColumnWidth(i, commonCellPropHolder.getColumnWidth());
                }
                if (endingTotalMapping.containsKey(i)){
                    if (calculateColumnIndexes.contains(i) || (calculateColumnIndexes.size() == 1 && calculateColumnIndexes.contains(-1))){
                        BigDecimal bigDecimal = endingTotalMapping.get(i);
                        Object convert = dataInverter.convert(bigDecimal);
                        cellValue = convert.toString();
                    }
                }
                cell.setCellValue(cellValue);

                if(cellProperty == null){
                    //未配置 合计部分样式继承上一行
                    CellStyle cellStyle = sheet.getRow(sheet.getLastRowNum() - 1).getCell(i).getCellStyle();
                    SXSSFWorkbook workbook = sheet.getWorkbook();
                    CellStyle totalCellStyle = workbook.createCellStyle();
                    StyleHelper.setCellStyleAlignment(totalCellStyle,cellStyle.getAlignment(),cellStyle.getVerticalAlignment());
                    totalCellStyle.setFillForegroundColor(cellStyle.getFillForegroundColorColor());
                    totalCellStyle.setFillPattern(cellStyle.getFillPattern());
                    totalCellStyle.setFont(workbook.getFontAt(cellStyle.getFontIndex()));

                    totalCellStyle.setBorderBottom(cellStyle.getBorderBottom());
                    totalCellStyle.setBorderLeft(cellStyle.getBorderLeft());
                    totalCellStyle.setBorderRight(cellStyle.getBorderRight());
                    totalCellStyle.setBorderTop(cellStyle.getBorderTop());
                    totalCellStyle.setLeftBorderColor(cellStyle.getLeftBorderColor());
                    totalCellStyle.setRightBorderColor(cellStyle.getRightBorderColor());
                    totalCellStyle.setTopBorderColor(cellStyle.getTopBorderColor());
                    totalCellStyle.setBottomBorderColor(cellStyle.getBottomBorderColor());
                    totalCellStyle.setDataFormat(StyleHelper.DATA_FORMAT_PLAIN_TEXT_INDEX);
                    cell.setCellStyle(totalCellStyle);
                }else{
                    cell.setCellStyle(createCellStyle(cellProperty));
                }

            }
        }
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH)){
            debug(LOGGER,"开始自动计算列宽");
            for (int columnIdx = 0; columnIdx < alreadyWrittenColumns; columnIdx++) {
                sheet.trackAllColumnsForAutoSizing();
                sheet.autoSizeColumn(columnIdx,true);
                sheet.setColumnWidth(columnIdx, (int) (sheet.getColumnWidth(columnIdx) * 1.35));
            }
        }
        Map<Integer, Integer> specialRowHeightMapping = writeConfig.getSpecialRowHeightMapping(sheetIndex);
        for (Map.Entry<Integer, Integer> heightEntry : specialRowHeightMapping.entrySet()) {
            debug(LOGGER,"设置工作表[%s]第%s行高度为%s",sheet.getSheetName(),heightEntry.getKey(),heightEntry.getValue());
            sheet.getRow(heightEntry.getKey()).setHeightInPoints(heightEntry.getValue());
        }
        if (writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_HIDDEN_BLANK_COLUMNS)){
            warn(LOGGER,"即将隐藏空白列，增加导出耗时");
            int maxColumnIndex = -1;
            for (Row cells : sheet) {
                maxColumnIndex = Math.max(cells.getLastCellNum(),maxColumnIndex);
            }
            if (maxColumnIndex > -1){
                for (int i = maxColumnIndex; i < SpreadsheetVersion.EXCEL2007.getMaxColumns(); i++) {
                    sheet.setColumnHidden(i,true);
                }
            }
        }
        return new AxolotlWriteResult(true, "完成结束阶段");
    }


    /**
     * 填充空白单元格
     * @param sheet 工作表
     */
    private void fillWhiteCell(Sheet sheet){
        CellStyle defaultStyle = createCellStyle(globalCellPropHolder);
        // 将默认样式应用到所有单元格
        for (int i = 0; i < 26; i++) {
            sheet.setDefaultColumnStyle(i, defaultStyle);
            sheet.setDefaultColumnWidth(globalCellPropHolder.getColumnWidth() == null ? DEFAULT_COLUMN_WIDTH : globalCellPropHolder.getColumnWidth());
        }
        sheet.setDefaultRowHeight(globalCellPropHolder.getRowHeight() == null ? DATA_ROW_HEIGHT : globalCellPropHolder.getRowHeight());
        debug(LOGGER,"渲染空白单元格");
    }

    /**
     * Part.1 表头
     * Step.1 创建标题行
     * @param sheet 工作表
     * @return 渲染结果
     */
    public AxolotlWriteResult createTitleRow(SXSSFSheet sheet){
        String title = writeConfig.getTitle();
        if (Validator.strNotBlank(title)){
            debug(LOGGER,"设置工作表标题:[%s]",title);
            int switchSheetIndex = context.getSwitchSheetIndex();
            Map<Integer, Integer> alreadyWriteRowMap = context.getAlreadyWriteRow();
            int alreadyWriteRow = alreadyWriteRowMap.getOrDefault(switchSheetIndex,-1);
            alreadyWriteRowMap.put(switchSheetIndex,++alreadyWriteRow);
            SXSSFRow titleRow = sheet.createRow(alreadyWriteRow);
            titleRow.setHeight(titleCellPropHolder.getRowHeight());
            SXSSFCell startPositionCell = titleRow.createCell(START_POSITION);
            startPositionCell.setCellValue(writeConfig.getTitle());
            return new AxolotlWriteResult(true, LoggerHelper.format("设置工作表标题:[%s]",title));
        }else{
            String message = "未设置工作表标题";
            debug(LOGGER,message);
            return new AxolotlWriteResult(false, message);
        }
    }


    /**
     * Part.1 表头
     * Step.2 递归表头
     * @param sheet 工作表
     */
    private AxolotlWriteResult defaultRenderHeaders(SXSSFSheet sheet){
        int switchSheetIndex = context.getSwitchSheetIndex();
        List<Header> headers = context.getHeaders().get(switchSheetIndex);
        int headerMaxDepth;
        int headerColumnCount = 0;
        int alreadyWriteRow = context.getAlreadyWriteRow().getOrDefault(context.getSwitchSheetIndex(),-1);
        if (headers != null && !headers.isEmpty()){
            List<Header> cacheHeaders;
            if (writeConfig.getWritePolicyAsBoolean(AUTO_INSERT_SERIAL_NUMBER)){
                CellPropertyHolder cellProperty = commonCellPropHolder.get(AUTO_INSERT_SERIAL_NUMBER);
                cacheHeaders = new ArrayList<>();
                if(cellProperty == null){
                    //未配置 走表头默认样式
                    cacheHeaders.add(new Header("序号"));
                }else{
                    //配置 走程序写入单元格的列宽
                    Header header = new Header("序号");
                    header.setCustomCellStyle(createCellStyle(cellProperty));
                    header.setColumnWidth(cellProperty.getColumnWidth());
                    cacheHeaders.add(header);
                }
                cacheHeaders.addAll(headers);
            }else{
                cacheHeaders = headers;
            }
            context.getAlreadyWriteRow().put(switchSheetIndex,++alreadyWriteRow);
            headerMaxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            debug(LOGGER,"起始行次为[%s]，表头最大深度为[%s]",alreadyWriteRow,headerMaxDepth);
            int sheetIndex = context.getSwitchSheetIndex();
            Map<String, Integer> headerCache = context.getHeaderColumnIndexMapping().row(sheetIndex);
            //根节点渲染
            for (Header header : cacheHeaders) {
                CellStyle usedCellStyle = getCellStyle(header, headerCellPropHolder);
                Row row = ExcelToolkit.createOrCatchRow(sheet, alreadyWriteRow);
                row.setHeight(headerCellPropHolder.getRowHeight());

                Cell cell = row.createCell(headerColumnCount, CellType.STRING);
                String title = header.getName();
                cell.setCellValue(title);
                int orlopCellNumber = header.countOrlopCellNumber();
                context.getAlreadyWrittenColumns().put(switchSheetIndex,context.getAlreadyWrittenColumns().getOrDefault(switchSheetIndex,0)+orlopCellNumber);
                debug(LOGGER,"渲染表头[%s],行[%s],列[%s],子表头列数量[%s]",title, alreadyWriteRow,headerColumnCount,orlopCellNumber);
                // 有子节点说明需要向下迭代并合并
                CellRangeAddress cellAddresses;
                if (header.getChilds()!=null && !header.getChilds().isEmpty()){
                    List<Header> childs = header.getChilds();
                    int childMaxDepth = ExcelToolkit.getMaxDepth(childs, 0);
                    cellAddresses = new CellRangeAddress(alreadyWriteRow, alreadyWriteRow +(headerMaxDepth-childMaxDepth)-1, headerColumnCount, headerColumnCount+orlopCellNumber-1);
                    HeaderRecursiveInfo headerRecursiveInfo = new HeaderRecursiveInfo();
                    headerRecursiveInfo.setAllRow(alreadyWriteRow+headerMaxDepth+1);
                    headerRecursiveInfo.setStartColumn(headerColumnCount);
                    headerRecursiveInfo.setAlreadyWriteColumn(headerColumnCount);
                    headerRecursiveInfo.setCellProperty(headerCellPropHolder);
                    headerRecursiveInfo.setRowHeight(headerCellPropHolder.getRowHeight());
                    recursionRenderHeaders(sheet,childs, headerRecursiveInfo);
                }else{
                    cellAddresses = new CellRangeAddress(alreadyWriteRow, (alreadyWriteRow +headerMaxDepth)-1, headerColumnCount, headerColumnCount);

                    String fieldName = header.getFieldName();
                    if (fieldName != null){
                        debug(LOGGER,"映射字段[%s]到列索引[%s]",fieldName,headerColumnCount);
                        headerCache.put(fieldName,headerColumnCount);
                    }
                    if (!writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH)){
                        int columnWidth = header.getColumnWidth();
                        if (columnWidth < 0){
                            if(headerCellPropHolder.getColumnWidth() != null){
                                columnWidth = headerCellPropHolder.getColumnWidth();
                            }else{
                                columnWidth = StyleHelper.getPresetCellLength(title);
                            }
                        }
                        debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",headerColumnCount,header.getName(),columnWidth);
                        sheet.setColumnWidth(headerColumnCount, columnWidth);
                    }else{
                        debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",headerColumnCount,header.getName(),"AUTO");
                    }
                    if (header.isParticipateInCalculate()){
                        debug(LOGGER,"列[%s]表头[%s]参与计算",headerColumnCount,header.getName());
                        writeConfig.addCalculateColumnIndex(sheetIndex,headerColumnCount);
                    }
                }
                StyleHelper.renderMergeRegionStyle(sheet,cellAddresses,usedCellStyle);
                if (headerMaxDepth > 1){
                    sheet.addMergedRegion(cellAddresses);
                }
                headerColumnCount+=orlopCellNumber;
            }
        }else{
            headerMaxDepth = 0;
            debug(LOGGER,"未设置表头");
        }
        context.getHeaderRowCount().put(switchSheetIndex,headerMaxDepth);
        alreadyWriteRow+=(headerMaxDepth-1);
        context.getAlreadyWriteRow().put(switchSheetIndex,alreadyWriteRow);
        context.getAlreadyWrittenColumns().put(switchSheetIndex,headerColumnCount);
        return new AxolotlWriteResult(true, "渲染表头成功");
    }


    @Data
    public class HeaderRecursiveInfo implements Serializable,Cloneable{

        /**
         * 渲染的总行数
         */
        private int allRow;

        /**
         * 起始列
         */
        private int startColumn;

        /**
         * 已经写入的列
         */
        private int alreadyWriteColumn;

        /**
         * 单元格属性
         */
        private CellPropertyHolder cellProperty;

        /**
         * 渲染的行高度
         */
        private short rowHeight;
    }

    /**
     * 递归渲染表头
     * @param sheet 工作表
     * @param headers 表头集合
     * @param headerRecursiveInfo 递归信息
     */
    @SneakyThrows
    private void recursionRenderHeaders(SXSSFSheet sheet, List<Header> headers, HeaderRecursiveInfo headerRecursiveInfo){
        if (headers != null && !headers.isEmpty()){
            int maxDepth = ExcelToolkit.getMaxDepth(headers, 0);
            int startRow = headerRecursiveInfo.getAllRow() - maxDepth -1;
            Row row = ExcelToolkit.createOrCatchRow(sheet,startRow);
            row.setHeight(headerRecursiveInfo.getRowHeight());
            int sheetIndex = context.getSwitchSheetIndex();
            Map<String, Integer> headerCache = context.getHeaderColumnIndexMapping().row(sheetIndex);
            for (Header header : headers) {
                CellStyle usedCellStyle = getCellStyle(header,headerRecursiveInfo.getCellProperty());
                int alreadyWriteColumn = headerRecursiveInfo.getAlreadyWriteColumn();
                Cell cell = ExcelToolkit.createOrCatchCell(sheet, row.getRowNum(), alreadyWriteColumn, null);
                cell.setCellValue(header.getName());
                int childCount = header.countOrlopCellNumber();
                int endColumnPosition = (alreadyWriteColumn + childCount);
                CellRangeAddress cellAddresses;
                int mergeRowNumber = startRow + maxDepth - 1;
                if (header.getChilds()!=null && !header.getChilds().isEmpty()){
                    cellAddresses = new CellRangeAddress(startRow, startRow, alreadyWriteColumn, endColumnPosition-1);
                }else{
                    cellAddresses = new CellRangeAddress(startRow, startRow + maxDepth-1, alreadyWriteColumn, endColumnPosition-1);

                    if (!writeConfig.getWritePolicyAsBoolean(ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH)){
                        int columnWidth = header.getColumnWidth();
                        if (columnWidth == -1){
                            columnWidth = StyleHelper.getPresetCellLength(header.getName());
                        }
                        debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",alreadyWriteColumn,header.getName(),columnWidth);
                        sheet.setColumnWidth(alreadyWriteColumn, columnWidth);
                    }else{
                        debug(LOGGER,"列[%s]表头[%s]设置列宽[%s]",alreadyWriteColumn,header.getName(),"AUTO");
                    }
                }
                StyleHelper.renderMergeRegionStyle(sheet,cellAddresses, usedCellStyle);
                if (mergeRowNumber !=  startRow){
                    sheet.addMergedRegion(cellAddresses);
                }
                headerRecursiveInfo.setAlreadyWriteColumn(endColumnPosition);
                headerRecursiveInfo.setStartColumn(alreadyWriteColumn);
                if (header.getChilds() != null && !header.getChilds().isEmpty()){
                    HeaderRecursiveInfo child = new HeaderRecursiveInfo();
                    BeanUtils.copyProperties(child, headerRecursiveInfo);
                    child.setAlreadyWriteColumn(headerRecursiveInfo.getStartColumn());
                    recursionRenderHeaders(sheet,header.getChilds(),child);
                }else{
                    String fieldName = header.getFieldName();
                    if (fieldName != null){
                        debug(LOGGER,"映射字段[%s]到列索引[%s]",fieldName,alreadyWriteColumn);
                        headerCache.put(fieldName,alreadyWriteColumn);
                    }
                    if (header.isParticipateInCalculate()){
                        debug(LOGGER,"列[%s]表头[%s]参与计算",alreadyWriteColumn,header.getName());
                        writeConfig.addCalculateColumnIndex(sheetIndex,alreadyWriteColumn);
                    }
                }
            }
        }
    }

    /**
     * Part.1 表头
     * 辅助方法 获取表头Header样式
     * @param header 表头
     * @param defaultCellStyle 表头样式默认配置
     * @return 表头样式
     */
    private CellStyle getCellStyle(Header header, CellPropertyHolder defaultCellStyle) {
        if (header.getCustomCellStyle() != null){
            debug(LOGGER,"["+header.getName()+"] Header内部样式配置读取完毕,优先级:高");
            return header.getCustomCellStyle();
        }else{
            AxolotlCellStyle axolotlCellStyle = header.getAxolotlCellStyle();
            if (axolotlCellStyle != null){
                CellPropertyHolder headerStyle = ConfigurableStyleConfig.copyConfigFromDefault(defaultCellStyle);
                if(axolotlCellStyle.getBorderLeftStyle() != null){
                    headerStyle.setBorderLeftStyle(axolotlCellStyle.getBorderLeftStyle());
                }
                if(axolotlCellStyle.getBorderRightStyle() != null){
                    headerStyle.setBorderRightStyle(axolotlCellStyle.getBorderRightStyle());
                }
                if(axolotlCellStyle.getBorderTopStyle() != null){
                    headerStyle.setBorderTopStyle(axolotlCellStyle.getBorderTopStyle());
                }
                if(axolotlCellStyle.getBorderBottomStyle() != null){
                    headerStyle.setBorderBottomStyle(axolotlCellStyle.getBorderBottomStyle());
                }
                if(axolotlCellStyle.getLeftBorderColor() != null){
                    headerStyle.setLeftBorderColor(axolotlCellStyle.getLeftBorderColor());
                }
                if(axolotlCellStyle.getRightBorderColor() != null){
                    headerStyle.setRightBorderColor(axolotlCellStyle.getRightBorderColor());
                }
                if(axolotlCellStyle.getTopBorderColor() != null){
                    headerStyle.setTopBorderColor(axolotlCellStyle.getTopBorderColor());
                }
                if(axolotlCellStyle.getBottomBorderColor() != null){
                    headerStyle.setBottomBorderColor(axolotlCellStyle.getBottomBorderColor());
                }
                if(axolotlCellStyle.getForegroundColor() != null){
                    headerStyle.setForegroundColor(axolotlCellStyle.getForegroundColor());
                }
                if(axolotlCellStyle.getFillPatternType() != null){
                    headerStyle.setFillPatternType(axolotlCellStyle.getFillPatternType());
                }
                if(axolotlCellStyle.getFontName() != null){
                    headerStyle.setFontName(axolotlCellStyle.getFontName());
                }
                if(axolotlCellStyle.getFontSize() != null){
                    headerStyle.setFontSize(axolotlCellStyle.getFontSize());
                }
                if(axolotlCellStyle.getFontColor() != null){
                    headerStyle.setFontColor(axolotlCellStyle.getFontColor());
                }
                if(axolotlCellStyle.getFontBold() != null){
                    headerStyle.setBold(axolotlCellStyle.getFontBold());
                }
                if(axolotlCellStyle.getItalic() != null){
                    headerStyle.setItalic(axolotlCellStyle.getItalic());
                }
                if(axolotlCellStyle.getStrikeout() != null){
                    headerStyle.setStrikeout(axolotlCellStyle.getStrikeout());
                }
                if(axolotlCellStyle.getHorizontalAlignment() != null){
                    headerStyle.setHorizontalAlignment(axolotlCellStyle.getHorizontalAlignment());
                }
                if(axolotlCellStyle.getVerticalAlignment() != null){
                    headerStyle.setVerticalAlignment(axolotlCellStyle.getVerticalAlignment());
                }
                debug(LOGGER,"["+header.getName()+"] Header内部样式配置读取完毕,优先级:中");
                return createCellStyle(headerStyle);
            }else{
                debug(LOGGER,"["+header.getName()+"] 未在Header内设置样式，选择默认表头样式配置");
                return createCellStyle(defaultCellStyle);
            }
        }
    }

    /**
     * 创建样式、从缓存中获取样式
     * @return 样式
     */
    private CellStyle createCellStyle(CellPropertyHolder cellProperty){
        CellPropertyHolder propertyHolder = ConfigurableStyleConfig.copyConfigFromDefault(cellProperty);
        propertyHolder.setColumnWidth(null);
        propertyHolder.setRowHeight(null);
        if(cellStyleCache.containsKey(propertyHolder)){
            return cellStyleCache.get(propertyHolder);
        }else{
            CellStyle style = createStyle(
                    propertyHolder.getBaseBorderStyle(),
                    propertyHolder.getBaseBorderColor(),
                    propertyHolder.getForegroundColor(),
                    propertyHolder.getFontName(),
                    propertyHolder.getFontSize(),
                    propertyHolder.getBold(),
                    propertyHolder.getFontColor(),
                    propertyHolder.getItalic(),
                    propertyHolder.getStrikeout()
            );
            //设置单元格对齐方式
            StyleHelper.setCellStyleAlignment(style, propertyHolder.getHorizontalAlignment(), propertyHolder.getVerticalAlignment());
            //设置单元格边框样式
            StyleHelper.setBorderStyle(style,propertyHolder);
            //填充样式
            style.setFillPattern(propertyHolder.getFillPatternType());
            cellStyleCache.put(propertyHolder,style);
            return style;
        }
    }

}

package cn.toutatis.xvoid.axolotl.excel.reader;

import cn.toutatis.xvoid.axolotl.Meta;
import cn.toutatis.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.toutatis.xvoid.axolotl.excel.reader.support.AxolotlAbstractExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException.ExceptionType;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkitKt;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class AxolotlExcelReader<T> extends AxolotlAbstractExcelReader<T> implements Iterator<List<T>> {

    /**
     * 日志工具
     */
    private final Logger LOGGER = LoggerToolkit.getLogger(AxolotlExcelReader.class);

    /**
     * 当前读取批次
     */
    private int currentReadBatch = -1;

    public AxolotlExcelReader(File excelFile) {
        super(excelFile);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(File excelFile, boolean withDefaultConfig) {
        super(excelFile, withDefaultConfig);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(File excelFile, Class<T> clazz) {
        super(excelFile, clazz);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(InputStream ins) {
        super(ins);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(InputStream ins, Class<T> clazz) {
        super(ins, clazz);
        super.LOGGER = LOGGER;
    }

    public AxolotlExcelReader(File excelFile, Class<T> clazz, boolean withDefaultConfig) {
        super(excelFile, clazz, withDefaultConfig);
        super.LOGGER = LOGGER;
    }

    /**
     * 是否有下一批数据
     */
    @Override
    public boolean hasNext() {
        return currentReadBatch * AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE < getRecordRowNumber();
    }

    /**
     * 获取下一批数据
     */
    @Override
    public List<T> next() {
        if (!hasNext()){
            throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR,"读取数据错误");
        }
        currentReadBatch++;
        LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME,"读取数据行数:"+currentReadBatch*AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE);
        return this.readSheetData(
                currentReadBatch*AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE,
                (currentReadBatch+1)*AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE
        );
    }

    /**
     * 读取表数据
     * 无任何形参，读取表中全部数据
     */
    public List<T> readSheetData(){
        return this.readSheetData(0);
    }

    /**
     * 读取表数据
     * @param start 起始位置
     */
    public List<T> readSheetData(int start){
        return this.readSheetData(
                _sheetLevelReaderConfig.getSheetName(),_sheetLevelReaderConfig.getSheetIndex(),
                0, this.getRecordRowNumber(),start
        );
    }

    /**
     * 读取起始和结束位置数据
     * 可以指定开始结束位置
     * @param start 起始位置
     * @param end 结束位置
     */
    public List<T> readSheetData(int start,int end){
        return this.readSheetData(
                _sheetLevelReaderConfig.getSheetName(),_sheetLevelReaderConfig.getSheetIndex(),
                start, end,0
        );
    }
    /**
     * 读取表数据
     * 可以指定开始结束位置和起始偏移行数
     * @param start 起始位置
     * @param end 结束位置
     * @param initialRowPositionOffset 初始行偏移量
     */
    public List<T> readSheetData(int start,int end,int initialRowPositionOffset){
        return this.readSheetData(
                _sheetLevelReaderConfig.getSheetName(),_sheetLevelReaderConfig.getSheetIndex(),
                start, end,initialRowPositionOffset
        );
    }

    /**
     * 使用表级配置读取数据
     *
     * @param sheetName 工作表名称
     * @param sheetIndex sheet索引
     * @param initialRowPositionOffset 初始行偏移量
     * @return 读取的数据
     */
    protected List<T> readSheetData(String sheetName,int sheetIndex,int start,int end,int initialRowPositionOffset){
        _sheetLevelReaderConfig.setSheetName(sheetName);
        _sheetLevelReaderConfig.setSheetIndex(sheetIndex);
        _sheetLevelReaderConfig.setStartIndex(start);
        _sheetLevelReaderConfig.setEndIndex(end);
        _sheetLevelReaderConfig.setInitialRowPositionOffset(initialRowPositionOffset);
        return this.readSheetData(_sheetLevelReaderConfig);
    }

    /**
     * @param castClass 读取的Java类型
     * @param sheetName 工作表名称
     */
    public <RT> List<RT> readSheetData(Class<RT> castClass,String sheetName){
        ReadConfigBuilder<RT> configBuilder = new ReadConfigBuilder<>(castClass, true);
        configBuilder.setSheetName(sheetName);
        return this.readSheetData(configBuilder);
    }

    /**
     * @param castClass 读取的Java类型
     * @param sheetIndex 表索引
     */
    public <RT> List<RT> readSheetData(Class<RT> castClass,int sheetIndex){
        ReadConfigBuilder<RT> configBuilder = new ReadConfigBuilder<>(castClass, true);
        configBuilder.setSheetIndex(sheetIndex);
        return this.readSheetData(configBuilder);
    }

    /**
     * @param castClass 读取的Java类型
     * @param initialRowPositionOffset 起始偏移量
     */
    public <RT> List<RT> readSheetDataOffset(Class<RT> castClass,int initialRowPositionOffset){
        ReadConfigBuilder<RT> configBuilder = new ReadConfigBuilder<>(castClass, true);
        configBuilder.setInitialRowPositionOffset(initialRowPositionOffset);
        return this.readSheetData(configBuilder);
    }

    /**
     * @param castClass 读取的Java类型
     */
    public <RT> List<RT> readSheetData(Class<RT> castClass){
        ReadConfigBuilder<RT> configBuilder = new ReadConfigBuilder<>(castClass, true);
        return this.readSheetData(configBuilder);
    }

    /**
     * 读取指定sheet的数据
     *
     * @param castClass 读取的类型
     * @param sheetIndex sheet索引
     * @param withDefaultConfig 是否使用默认配置
     * @param startIndex 起始行
     * @param endIndex 结束行
     * @param initialRowPositionOffset 起始行偏移量
     * @param <RT>  类型泛型
     * @return 读取的数据
     */
    public <RT> List<RT> readSheetData(Class<RT> castClass,int sheetIndex,boolean withDefaultConfig,
                                       int startIndex,int endIndex,int initialRowPositionOffset) {
        ReadConfigBuilder<RT> configBuilder = new ReadConfigBuilder<>(castClass, withDefaultConfig);
        configBuilder
                .setSheetIndex(sheetIndex)
                .setStartIndex(startIndex)
                .setEndIndex(endIndex)
                .setInitialRowPositionOffset(initialRowPositionOffset);
        return this.readSheetData(configBuilder);
    }

    /**
     * 使用读取配置构建读取配置
     *
     * @param configBuilder 读取配置构建器
     * @return 读取数据
     * @param <RT>  读取的类型泛型
     */
    public <RT> List<RT> readSheetData(ReadConfigBuilder<RT> configBuilder) {
        return this.readSheetData(configBuilder.build());
    }

    /**
     * [ROOT]
     * 读取Excel数据
     *
     * @param readerConfig 读取配置
     * @return 读取数据
     * @param <RT>  读取的类型泛型
     */
    public <RT> List<RT> readSheetData(ReaderConfig<RT> readerConfig) {
        List<RT> readResult = new ArrayList<>();
        // 查找sheet
        Sheet sheet = this.searchSheet(readerConfig);
        // 检查并修正配置文件
        this.preCheckAndFixReadConfig(readerConfig);
        // 空表返回空list
        if (sheet == null){
            return readResult;
        }
        // 处理合并单元格
        this.spreadMergedCells(sheet);
        this.readSheetData(sheet,readerConfig,readResult);
        return readResult;
    }

    /**
     * @param readerConfig 读取配置
     */
    private Sheet searchSheet(ReaderConfig<?> readerConfig){
        if (readerConfig == null){return null;}
        Sheet sheet;
        if (readerConfig.getSheetName() != null){
            sheet = workBookContext.getWorkbook().getSheet(readerConfig.getSheetName());
            if (sheet != null){
                readerConfig.setSheetIndex(sheet.getWorkbook().getSheetIndex(sheet));
            }else {
                readerConfig.setSheetIndex(-1);
            }
        }else {
            try {
                sheet = workBookContext.getIndexSheet(readerConfig.getSheetIndex());
            }catch (IllegalArgumentException e){
                if (e.getMessage().contains("out of range")){
                    int numberOfSheets = workBookContext.getWorkbook().getNumberOfSheets()-1;
                    ;
                    LoggerToolkitKt.warnWithModule(LOGGER, Meta.MODULE_NAME,
                            String.format("表索引[%s]超出范围[0-%s],将返回空数据或抛出异常",readerConfig.getSheetIndex(),numberOfSheets)
                    );
                }else {
                    throw e;
                }
                sheet = null;
                readerConfig.setSheetIndex(-1);
            }
        }
        return sheet;
    }

    /**
     * 处理合并单元格
     *
     * @param sheet 工作表
     */
    private void spreadMergedCells(Sheet sheet) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME, "开始处理工作表合并单元格");
        for (CellRangeAddress mergedRegion : mergedRegions) {
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = mergedRegion.getLastRow();
            int firstColumn = mergedRegion.getFirstColumn();
            int lastColumn = mergedRegion.getLastColumn();
            Cell leftTopCell = sheet.getRow(firstRow).getCell(firstColumn);
            LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME, String.format("处理合并单元格[%s]",mergedRegion.formatAsString()));
            for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
                for (int columnIndex = firstColumn; columnIndex <= lastColumn; columnIndex++) {
                    Row row = sheet.getRow(rowIndex);
                    Cell cell =row.getCell(columnIndex);
                    if (cell == null){
                        cell = row.createCell(columnIndex, leftTopCell.getCellType());
                    }
                    switch (leftTopCell.getCellType()){
                        case STRING:
                            cell.setCellValue(leftTopCell.getStringCellValue());
                            break;
                        case NUMERIC:
                            cell.setCellValue(leftTopCell.getNumericCellValue());
                            break;
                        case BOOLEAN:
                            cell.setCellValue(leftTopCell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            cell.setCellValue(leftTopCell.getCellFormula());
                    }
                }
            }
        }
    }

    /**
     * 读取表中每一行的数据
     *
     * @param readerConfig 读取配置
     */
    private <RT> void readSheetData(Sheet sheet,ReaderConfig<RT> readerConfig,List<RT> list){
        int startIndex = readerConfig.getStartIndex();
        int endIndex = readerConfig.getEndIndex();
        if (startIndex == 0){
            int initialRowPositionOffset = readerConfig.getInitialRowPositionOffset();
            if (initialRowPositionOffset > 0){
                LOGGER.debug("跳过前{}行",initialRowPositionOffset);
                startIndex = startIndex + initialRowPositionOffset;
            }
        }
        this.searchHeaderCellPosition(readerConfig);
        for (int i = startIndex; i < endIndex; i++) {
            RT instance = this.readRow(sheet, i, readerConfig);
            if (instance!= null){list.add(instance);}
        }
    }

    /**
     * 查找表头
     *
     * @param readerConfig 读取配置
     */
    private void matchHeaderCellPosition(ReaderConfig<?> readerConfig){
        int readHeadRows = Math.min(getRecordRowNumber(readerConfig), AxolotlDefaultReaderConfig.XVOID_DEFAULT_HEADER_FINDING_ROW);
        List<EntityCellMappingInfo<?>> indexMappingInfos = readerConfig.getIndexMappingInfos();
        // 提取表头
        Map<String, Integer> headerKeys = indexMappingInfos.stream()
                .map(EntityCellMappingInfo::getHeaderName)
                .filter(Objects::nonNull)
                .distinct()
                .collect(Collectors.toMap(element -> element, i -> -1));
        if (!headerKeys.isEmpty()){
            Map<Integer, HashBasedTable<String, Integer, Integer>> headerCaches = workBookContext.getHeaderCaches();
            // 表头名称,序号,列位置
            HashBasedTable<String, Integer, Integer> headerCache;
            boolean hintCache = false;
            if (headerCaches.containsKey(readerConfig.getSheetIndex())){
                LoggerHelper.debug(LOGGER,LoggerHelper.format("从缓存中获取表头,数量[%s]", headerKeys.size()));
                headerCache = headerCaches.get(readerConfig.getSheetIndex());
                hintCache = true;
            }else {
                LoggerHelper.debug(LOGGER,LoggerHelper.format("开始查找表头,数量[%s],查找表头:%s", headerKeys.size(),headerKeys));
                headerCache = HashBasedTable.create();
            }
            Sheet sheet = workBookContext.getIndexSheet(readerConfig.getSheetIndex());
            for (int i = 0; i < readHeadRows; i++) {
                Row row = sheet.getRow(i);
                if (ExcelToolkit.notBlankRowCheck(row)){
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()){
                        Cell cell = cellIterator.next();
                        if (cell != null && cell.getCellType() == CellType.STRING){
                            String cellValue = cell.getStringCellValue();
                            if (headerKeys.containsKey(cellValue)){
                                LoggerHelper.debug(LOGGER,LoggerHelper.format("查找到表头[%s]", cellValue));
                                headerCache.put(cellValue, headerCache.row(cellValue).size()+1 , cell.getColumnIndex());
                            }
                        }
                    }
                }
            }
            LoggerHelper.debug(LOGGER,LoggerHelper.format("查找表头结束,映射信息:%s", headerCache));
            if (!hintCache){
                headerCaches.put(readerConfig.getSheetIndex(), headerCache);
            }
            for (EntityCellMappingInfo<?> indexMappingInfo : indexMappingInfos) {
                String headerName = indexMappingInfo.getHeaderName();
                if (StringUtils.isNotBlank(headerName)){
                    Map<Integer, Integer> recordInfo = headerCache.row(headerName);
                    if (recordInfo.isEmpty()){
                        if (readerConfig.getReadPolicyAsBoolean(ReadPolicy.IGNORE_EMPTY_SHEET_HEADER_ERROR)){
                            LoggerHelper.debug(LOGGER,LoggerHelper.format("表头[%s]不存在", headerName));
                        }else {
                            throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR,LoggerHelper.format("表头[%s]不存在", headerName));
                        }
                        continue;
                    }
                    Integer assignedIndex = headerKeys.get(headerName);
                    Integer columnIndex;
                    int headerNameIndex = indexMappingInfo.getHeaderNameIndex();
                    if (headerNameIndex == -1){
                        if (assignedIndex == -1){
                            columnIndex = recordInfo.get(1);
                            assignedIndex = 1;
                        }else {
                            assignedIndex+=1;
                            columnIndex = recordInfo.get(assignedIndex);
                        }
                        LoggerHelper.debug(LOGGER,LoggerHelper.format("映射同名表头[%s]到列[%s]", headerName, columnIndex));
                        headerKeys.put(headerName, assignedIndex);
                    }else{
                        int idx = headerNameIndex + 1;
                        LoggerHelper.debug(LOGGER,LoggerHelper.format("指定同名表头[%s]列为[%s]", headerName, idx));
                        columnIndex = recordInfo.get(idx) != null ? recordInfo.get(idx) : -1;
                    }
                    indexMappingInfo.setColumnPosition(columnIndex);
                }
            }
        }
    }

    /**
     * [ROOT]
     * 读取行信息到对象
     *
     * @param sheet 表
     * @param rowNumber 行号
     * @param readerConfig 读取配置
     * @param <RT>  转换类型
     */
    private <RT> RT readRow(Sheet sheet,int rowNumber,ReaderConfig<RT> readerConfig){
        RT instance = readerConfig.getCastClassInstance();
        Row row = sheet.getRow(rowNumber);
        if (ExcelToolkit.blankRowCheck(row)){
            if (readerConfig.getReadPolicyAsBoolean(ReadPolicy.INCLUDE_EMPTY_ROW)){
                return instance;
            }else{
                return null;
            }
        }
        this.convertCellToInstance(row,instance,readerConfig);
        return instance;
    }

    /**
     * @param readerConfig 读取配置
     */
    @SuppressWarnings({"unchecked","rawtypes"})
    private <RT> void convertCellToInstance(Row row,RT instance,ReaderConfig<RT> readerConfig){
        if (instance instanceof Map){
            this.row2MapInstance((Map)instance, row,readerConfig);
        }else{
            this.row2SimplePOJO(instance,row,readerConfig);
        }
    }

    /**
     * 填充单元格数据到Java对象
     *
     * @param instance 对象实例
     * @param row 行
     * @param readerConfig 读取配置
     * @param <RT>  读取类型
     */
    @SneakyThrows
    private <RT> void row2SimplePOJO(RT instance, Row row, ReaderConfig<RT> readerConfig){
        this.convertPositionCellToInstance(instance,readerConfig,row.getSheet());
        List<EntityCellMappingInfo<?>> indexMappingInfos = readerConfig.getIndexMappingInfos();
        for (EntityCellMappingInfo<?> indexMappingInfo : indexMappingInfos) {
            workBookContext.setCurrentReadRowIndex(row.getRowNum());
            workBookContext.setCurrentReadColumnIndex(indexMappingInfo.getColumnPosition());
            CellGetInfo cellValue = this.getCellOriginalValue(row,indexMappingInfo.getColumnPosition(), indexMappingInfo);
            Object adaptiveValue = this.adaptiveCellValue2EntityClass(cellValue, indexMappingInfo, readerConfig);
            this.assignValueToField(instance,adaptiveValue,indexMappingInfo,readerConfig);
        }
        this.validateConvertEntity(instance, readerConfig.getReadPolicyAsBoolean(ReadPolicy.VALIDATE_READ_ROW_DATA));
    }

    /**
     * @param readerConfig 读取配置
     */
    private void convertPositionCellToInstance(Object instance,ReaderConfig<?> readerConfig,Sheet sheet){
        List<EntityCellMappingInfo<?>> positionMappingInfos = readerConfig.getPositionMappingInfos();
        for (EntityCellMappingInfo<?> positionMappingInfo : positionMappingInfos) {
            workBookContext.setCurrentReadRowIndex(positionMappingInfo.getRowPosition());
            workBookContext.setCurrentReadColumnIndex(positionMappingInfo.getColumnPosition());
            CellGetInfo cellValue = this.getPositionCellOriginalValue(sheet, positionMappingInfo);
            Object adaptiveValue = this.adaptiveCellValue2EntityClass(cellValue, positionMappingInfo, readerConfig);
            this.assignValueToField(instance,adaptiveValue,positionMappingInfo,readerConfig);
        }
    }

    /**
     * [ROOT]
     * 讲适配后的值赋值给实体类字段
     *
     * @param instance 实体类实例
     * @param adaptiveValue 适配后的值
     * @param mappingInfo 映射信息
     * @param readerConfig 读取配置
     */
    @SneakyThrows
    private void assignValueToField(Object instance,Object adaptiveValue, EntityCellMappingInfo<?> mappingInfo,ReaderConfig<?> readerConfig){
        Field field = instance.getClass().getDeclaredField(mappingInfo.getFieldName());
        field.setAccessible(true);
        Object o = field.get(instance);
        if (o!= null){
            if (readerConfig.getReadPolicyAsBoolean(ReadPolicy.FIELD_EXIST_OVERRIDE)){
                field.set(instance, adaptiveValue);
            }
        }else {
            field.set(instance, adaptiveValue);
        }
    }

    /**
     * [ROOT]
     * 适配实体类的字段
     *
     * @param info 单元格值
     * @param mappingInfo 映射信息
     * @param readerConfig 读取配置
     * @return 适配实体类的字段值
     */
    @SuppressWarnings({"unchecked","rawtypes"})
    private Object adaptiveCellValue2EntityClass(CellGetInfo info, EntityCellMappingInfo<?> mappingInfo, ReaderConfig<?> readerConfig){
        if (mappingInfo.getDataCastAdapter() == AutoAdapter.class){
            DataCastAdapter<Object> autoAdapter = AutoAdapter.instance();
            return this.adaptiveValue(autoAdapter,info, (EntityCellMappingInfo<Object>) mappingInfo, (ReaderConfig<Object>) readerConfig);
        }
        Class<? extends DataCastAdapter<?>> dataCastAdapterClass = mappingInfo.getDataCastAdapter();
        if (dataCastAdapterClass != null && !dataCastAdapterClass.isInterface()){
            DataCastAdapter adapter;
            try {
                adapter = dataCastAdapterClass.getDeclaredConstructor().newInstance();
                return this.adaptiveValue(adapter,info, (EntityCellMappingInfo<Object>) mappingInfo, (ReaderConfig<Object>) readerConfig);
            } catch (InstantiationException | IllegalAccessException |
                     InvocationTargetException | NoSuchMethodException e) {
                throw new AxolotlExcelReadException(ExceptionType.CONVERT_FIELD_ERROR,e);
            }
        }else {
            throw new AxolotlExcelReadException(mappingInfo,String.format("[%s]字段请配置适配器,字段类型:[%s]",mappingInfo.getFieldName(), mappingInfo.getFieldType()));
        }
    }

    /**
     * @param readerConfig 读取配置
     */
    private Object adaptiveValue(DataCastAdapter<Object> adapter, CellGetInfo info, EntityCellMappingInfo<Object> mappingInfo, ReaderConfig<Object> readerConfig) {
        if (adapter == null){
            throw new AxolotlExcelReadException(mappingInfo,String.format("未找到转换的类型:[%s->%s],字段:[%s]",info.getCellType(), mappingInfo.getFieldType(), mappingInfo.getFieldName()));
        }
        if (adapter instanceof AbstractDataCastAdapter<Object> abstractDataCastAdapter){
            abstractDataCastAdapter.setReaderConfig(readerConfig);
            abstractDataCastAdapter.setEntityCellMappingInfo(mappingInfo);
            return castValue(abstractDataCastAdapter, info, mappingInfo);
        }
        return castValue(adapter, info, mappingInfo);
    }

    /**
     *
     */
    private Object castValue(DataCastAdapter<Object> adapter, CellGetInfo info, EntityCellMappingInfo<Object> mappingInfo) {
        if (!adapter.support(info.getCellType(), mappingInfo.getFieldType())){
            throw new AxolotlExcelReadException(mappingInfo,String.format("不支持转换的类型:[%s->%s],字段:[%s]",info.getCellType(), mappingInfo.getFieldType(), mappingInfo.getFieldName()));
        }
        CastContext<Object> castContext = new CastContext<>(
                mappingInfo.getFieldType(), mappingInfo.getFormat(),
                workBookContext.getCurrentReadColumnIndex(), workBookContext.getCurrentReadRowIndex()
        );
        return adapter.cast(info, castContext);
    }

    /**
     * 获取位置映射单元格原始值
     *
     * @param sheet 表
     * @param mappingInfo 映射信息
     * @return 单元格值
     */
    private CellGetInfo getPositionCellOriginalValue(Sheet sheet, EntityCellMappingInfo<?> mappingInfo){
        Row row = sheet.getRow(mappingInfo.getRowPosition());
        if (row == null){
            return this.getBlankCellValue(mappingInfo);
        }
        Cell cell = row.getCell(mappingInfo.getColumnPosition());
        if (cell == null){
            return this.getBlankCellValue(mappingInfo);
        }
        return this.getCellOriginalValue(row,mappingInfo.getColumnPosition(), mappingInfo);
    }

    /**
     * 获取单元格原始值
     *
     * @param row 行次
     * @param mappingInfo 映射信息
     * @see #getIndexCellValue(Row, int, EntityCellMappingInfo)
     * @return 单元格值
     */
    private CellGetInfo getCellOriginalValue(Row row,int index, EntityCellMappingInfo<?> mappingInfo){
        // 一般不为null，由map类型传入时，默认使用索引映射
        if (mappingInfo == null){
            mappingInfo = new EntityCellMappingInfo<>(String.class);
            mappingInfo.setColumnPosition(index);
        }
        return this.getIndexCellValue(row,index, mappingInfo);
    }

    /**
     * 获取索引映射单元格值
     *
     * @param row 行次
     * @param mappingInfo 映射信息
     * @see #getBlankCellValue(EntityCellMappingInfo)
     * @return 单元格值
     * @see #getFormulaCellValue(Cell)
     */
    private CellGetInfo getIndexCellValue(Row row,int index, EntityCellMappingInfo<?> mappingInfo){
        if (index < 0){
            return this.getBlankCellValue(mappingInfo);
        }
        Cell cell = row.getCell(index);
        if (mappingInfo.getColumnPosition() == -1 || cell == null){
            return this.getBlankCellValue(mappingInfo);
        }
        Object value = null;
        CellGetInfo cellGetInfo = new CellGetInfo();
        CellType cellType = cell.getCellType();
        cellGetInfo.setCellType(cellType);
        switch (cellType) {
            case STRING:
                value = cell.getStringCellValue();
                break;
            case NUMERIC:
                cellGetInfo.set_cell(cell);
                value = cell.getNumericCellValue();
                break;
            case BOOLEAN:
                value = cell.getBooleanCellValue();
                break;
            case FORMULA:
                cellGetInfo = getFormulaCellValue(cell);
                return cellGetInfo;
            case BLANK:
                LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME,String.format("空白单元格位置:[%s]",workBookContext.getHumanReadablePosition()));
                return this.getBlankCellValue(mappingInfo);
            default:
                LOGGER.error(
                        "未知的单元格类型:{},单元格位置:[{}]", cell.getCellType(),
                        workBookContext.getHumanReadablePosition()
                );
                break;
        }
        cellGetInfo.setAlreadyFillValue(true);
        cellGetInfo.setCellValue(value);
        return cellGetInfo;
    }

    /**
     * 获取空单元格值
     *
     * @param mappingInfo 映射信息
     * @return 默认填充值
     */
    private CellGetInfo getBlankCellValue(EntityCellMappingInfo<?> mappingInfo){
        CellGetInfo cellGetInfo = new CellGetInfo();
        if (mappingInfo.fieldIsPrimitive()){
            cellGetInfo.setCellValue(mappingInfo.fillDefaultPrimitiveValue(null));
        }
        return cellGetInfo;
    }

    /**
     * [ROOT]
     * 填充单元格数据到map
     *
     * @param readerConfig 读取配置
     */
    private <RT> void row2MapInstance(Map<String,Object> instance, Row row,ReaderConfig<RT> readerConfig){
        workBookContext.setCurrentReadRowIndex(row.getRowNum());
        row.cellIterator().forEachRemaining(cell -> {
            workBookContext.setCurrentReadColumnIndex(cell.getColumnIndex());
            int idx = cell.getColumnIndex() + 1;
            String key = "CELL_" + idx;
            instance.put(key, getCellOriginalValue(row, cell.getColumnIndex(),null).getCellValue());
            if (readerConfig.getReadPolicyAsBoolean(ReadPolicy.USE_MAP_DEBUG)){
                instance.put("CELL_TYPE_"+idx,cell.getCellType());
                if (cell.getCellType() == CellType.NUMERIC){
                    if (DateUtil.isCellDateFormatted(cell)){
                        instance.put("CELL_TYPE_"+idx,cell.getCellType());
                        instance.put("CELL_DATE_"+idx, Time.regexTime(cell.getDateCellValue()));
                    }else{
                        instance.put("CELL_TYPE_"+idx,cell.getCellType());
                    }
                }else {
                    instance.put("CELL_TYPE_"+idx,cell.getCellType());
                }
            }
        });
    }

    /**
     * 校验读取实体是否符合验证规则
     */
    private <RT> void validateConvertEntity(RT instance, boolean isValidate) {
        if (isValidate){
            Set<ConstraintViolation<RT>> validate = validator.validate(instance);
            if (!validate.isEmpty()){
                for (ConstraintViolation<RT> constraintViolation : validate) {
                    throw new AxolotlExcelReadException(workBookContext, constraintViolation.getMessage());
                }
            }
        }
    }

    /**
     * [ROOT]
     * 预校验读取配置是否正常
     * 不正常的数据将被修正
     *
     * @param readerConfig 读取配置
     */
    private void preCheckAndFixReadConfig(ReaderConfig<?> readerConfig) {
        //检查部分
        if (readerConfig == null){
            String msg = "读取配置不能为空";
            LOGGER.error(msg);
            throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR,msg);
        }
        int sheetIndex = readerConfig.getSheetIndex();
        if (sheetIndex < 0){
            String msg = String.format("读取的sheet不存在[%s]",readerConfig.getSheetName() != null? readerConfig.getSheetName() : readerConfig.getSheetIndex());
            if (readerConfig.getReadPolicyAsBoolean(ReadPolicy.IGNORE_EMPTY_SHEET_ERROR)){
                LoggerToolkitKt.warnWithModule(LOGGER,Meta.MODULE_NAME,msg+"将返回空数据");
                return;
            }
            throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR,msg);
        }
        Class<?> castClass = readerConfig.getCastClass();
        if (castClass == null){
            throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR,"读取的类型对象不能为空");
        }
        if (readerConfig.getStartIndex() < 0){
            throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR,"读取起始行不得小于0");
        }
        if (readerConfig.isReadAsObject()){
            if (!readerConfig.getIndexMappingInfos().isEmpty()){
                String simpleName = ColumnBind.class.getSimpleName();
                LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME,"读取对象时不用指定@"+simpleName+"注解");
            }
        }
        //修正部分
        if (readerConfig.getInitialRowPositionOffset() < 0){
            LOGGER.warn("读取的初始行偏移量不能小于0，将被修正为0");
            readerConfig.setInitialRowPositionOffset(0);
        }
        if (readerConfig.getEndIndex() < 0){
            if (readerConfig.getEndIndex() != -1){
                throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR,"读取结束行不得小于0");
            }
            if (!readerConfig.isReadAsObject()){
                LOGGER.info("未设置读取的结束行,将被默认修正为读取该表最大行数");
            }
            readerConfig.setEndIndex(workBookContext.getIndexSheet(sheetIndex).getLastRowNum()+1);
        }
    }

    /**
     * [ROOT]
     * 计算单元格公式为结果
     *
     * @param cell 单元格
     * @return 计算结果
     */
    private CellGetInfo getFormulaCellValue(Cell cell) {
        // 从元数据中获取计算计算器
        CellValue evaluated = workBookContext.getFormulaEvaluator().evaluate(cell);
        // 将单元格为公式的单元格值转换为计算结果
        Object value;
        switch (evaluated.getCellType()) {
            case STRING:
                value = evaluated.getStringValue();
                break;
            case NUMERIC:
                value = evaluated.getNumberValue();
                break;
            case BOOLEAN:
                value = evaluated.getBooleanValue();
                break;
            default:
                String msg = String.format("未知的公式单元格类型位置:[%d,%d],单元格类型:[%s],单元格值:[%s]",
                        cell.getRowIndex(), cell.getColumnIndex(), evaluated.getCellType(), evaluated);
                LOGGER.error(msg);
                throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR, msg);
        }
        CellGetInfo cellGetInfo = new CellGetInfo(true, value);
        cellGetInfo.setCellType(evaluated.getCellType());
        return cellGetInfo;
    }

    /**
     *
     */
    public int getPhysicalRowNumber(){
        return getRowNumber(true);
    }

    /**
     *
     */
    public int getRecordRowNumber(){
        return getRowNumber(false);
    }

    /**
     * 获取行数
     *
     * @param isPhysical 是否是物理行数
     * @return 行数
     */
    public int getRowNumber(boolean isPhysical){
        return getRowNumber(_sheetLevelReaderConfig,isPhysical);
    }

    /**
     * 获取行数
     *
     * @param readerConfig 读取配置
     * @param isPhysical 是否是物理行数
     * @return 行数
     */
    public int getRowNumber(ReaderConfig<?> readerConfig,boolean isPhysical){
        return getRowNumber(readerConfig.getSheetIndex(),isPhysical);
    }

    /**
     * @param readerConfig 读取配置
     */
    public int getPhysicalRowNumber(ReaderConfig<?> readerConfig){
        return getRowNumber(readerConfig.getSheetIndex(),true);
    }

    /**
     * @param readerConfig 读取配置
     */
    public int getRecordRowNumber(ReaderConfig<?> readerConfig){
        return getRowNumber(readerConfig.getSheetIndex(),false);
    }

    /**
     * [ROOT]
     * 获取行数
     *
     * @param sheetIndex 表索引
     * @param isPhysical 是否是物理行数
     * @return 行数
     */
    public int getRowNumber(int sheetIndex,boolean isPhysical){
        Sheet sheet = workBookContext.getIndexSheet(sheetIndex);
        if (isPhysical){
            return sheet.getPhysicalNumberOfRows();
        }else {
            return sheet.getLastRowNum()+1;
        }
    }

    /**
     * 获取当前读取位置
     */
    public String getHumanReadablePosition(){
        return workBookContext.getHumanReadablePosition();
    }

    /**
     * 当前读取批次
     */
    private int currentReadBatch = -1;

    /**
     * 是否有下一批数据
     */
    @Override
    public boolean hasNext() {
        return currentReadBatch * AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE < getRecordRowNumber();
    }

    /**
     * 获取下一批数据
     */
    @Override
    public List<T> next() {
        if (!hasNext()){
            throw new AxolotlExcelReadException(ExceptionType.READ_EXCEL_ERROR,"读取数据错误");
        }
        currentReadBatch++;
        LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME,"读取数据行数:"+currentReadBatch*AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE);
        return this.readSheetData(
                currentReadBatch*AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE,
                (currentReadBatch+1)*AxolotlDefaultReaderConfig.XVOID_DEFAULT_READ_EACH_BATCH_SIZE
        );
    }
}

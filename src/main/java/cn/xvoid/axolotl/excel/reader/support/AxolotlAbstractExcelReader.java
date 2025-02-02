package cn.xvoid.axolotl.excel.reader.support;

import cn.xvoid.axolotl.Meta;
import cn.xvoid.axolotl.excel.reader.ReadConfigBuilder;
import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.WorkBookContext;
import cn.xvoid.axolotl.excel.reader.annotations.AxolotlReaderSetter;
import cn.xvoid.axolotl.excel.reader.annotations.ColumnBind;
import cn.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.adapters.AbstractDataCastAdapter;
import cn.xvoid.axolotl.excel.reader.support.adapters.AutoAdapter;
import cn.xvoid.axolotl.excel.reader.support.docker.AxolotlCellMapInfo;
import cn.xvoid.axolotl.excel.reader.support.docker.MapDocker;
import cn.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.xvoid.axolotl.excel.writer.style.ComponentRender;
import cn.xvoid.axolotl.toolkit.ExcelToolkit;
import cn.xvoid.axolotl.toolkit.LoggerHelper;
import cn.xvoid.axolotl.toolkit.tika.DetectResult;
import cn.xvoid.axolotl.toolkit.tika.TikaShell;
import cn.xvoid.common.standard.StringPool;
import cn.xvoid.toolkit.clazz.ClassToolkit;
import cn.xvoid.toolkit.clazz.ReflectToolkit;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.toolkit.log.LoggerToolkitKt;
import cn.xvoid.axolotl.excel.reader.annotations.SpecifyPositionBind;
import com.google.common.collect.HashBasedTable;
import com.google.common.io.ByteStreams;
import jakarta.validation.ConstraintViolation;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.ValidatorFactory;
import lombok.Getter;
import lombok.Setter;
import lombok.SneakyThrows;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.RecordFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.apache.tika.mime.MimeType;
import org.slf4j.Logger;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Modifier;
import java.text.DecimalFormat;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public abstract class AxolotlAbstractExcelReader<T> {

    /**
     * 日志工具
     */
    protected Logger LOGGER = LoggerToolkit.getLogger(Meta.MODULE_NAME);

    /**
     * 工作簿元信息
     */
    @Getter
    protected WorkBookContext workBookContext;

    /**
     * JSR-303校验器
     * 使用Hibernate实现的校验器完成对读取数据的验证
     */
    protected Validator validator;

    /**
     * [内部属性]
     * 直接指定读取的类
     * 在读取数据时使用不指定读取类型的读取方法时，使用该类读取数据
     */
    @Setter
    protected ReaderConfig<T> _sheetLevelReaderConfig;

    private final ComponentRender componentRender = new ComponentRender();

    /**
     * 构造文件读取器
     */
    public AxolotlAbstractExcelReader(File excelFile) {
        this(excelFile,true);
    }

    /**
     * 构造文件读取器
     * @param excelFile 工作簿文件
     * @param withDefaultConfig 是否使用默认配置
     */
    @SuppressWarnings("unchecked")
    public AxolotlAbstractExcelReader(File excelFile, boolean withDefaultConfig) {
        this(excelFile, (Class<T>) Object.class,withDefaultConfig);
    }

    /**
     * 构造文件读取器
     * @param excelFile 工作簿文件
     * @param clazz 读取的类
     */
    public AxolotlAbstractExcelReader(File excelFile, Class<T> clazz) {
        this(excelFile,clazz,true);
    }

    /**
     * [ROOT]
     * 流支持构造
     *
     * @param ins 文件流
     */
    @SuppressWarnings("unchecked")
    public AxolotlAbstractExcelReader(InputStream ins) {
        this(ins, (Class<T>) Object.class);
    }

    public AxolotlAbstractExcelReader(InputStream ins, Class<T> clazz) {
        if (ins == null){
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,"文件流为空");
        }
        ByteArrayOutputStream dataCacheOutputStream =  new ByteArrayOutputStream();
        try {
            ByteStreams.copy(ins, dataCacheOutputStream);
            ins.close();
        } catch (IOException e) {
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,e.getMessage());
        }
        DetectResult detectResult = this.checkFileFormat(null, new ByteArrayInputStream(dataCacheOutputStream.toByteArray()));
        this.workBookContext = new WorkBookContext(new ByteArrayInputStream(dataCacheOutputStream.toByteArray()),detectResult);
        this._sheetLevelReaderConfig = new ReaderConfig<>(clazz,true);
        this.loadFileDataToWorkBook();
        this.createAdditionalExtensions();
    }

    /**
     * [ROOT]
     * 构造文件读取器
     * 初始化读取Excel文件
     * 1.初始化加载文件先判断文件是否正常并且是需要的格式
     * 2.将文件加载到POI工作簿中
     *
     * @param excelFile Excel工作簿文件
     * @param withDefaultConfig 是否使用默认配置
     */
    public AxolotlAbstractExcelReader(File excelFile, Class<T> clazz, boolean withDefaultConfig) {
        if (clazz == null){
            throw new IllegalArgumentException("读取的类型对象不能为空");
        }
        DetectResult predCheckFileNormal = TikaShell.preCheckFileNormal(excelFile);
        if (!predCheckFileNormal.isDetect()){
            predCheckFileNormal.throwException();
        }
        DetectResult detectResult = this.checkFileFormat(excelFile, null);
        workBookContext = new WorkBookContext(excelFile,detectResult);
        this._sheetLevelReaderConfig = new ReaderConfig<>(clazz,withDefaultConfig);
        this.loadFileDataToWorkBook();
        this.createAdditionalExtensions();
        this.componentRender.setReader(true);
        this.componentRender.setConfig(this._sheetLevelReaderConfig);
    }

    /**
     * 获取工作表信息
     * <p>
     * 本方法从当前的工作簿中获取所有工作表的名称并返回一个工作表名称的列表
     *
     * @return 包含所有工作表名称的列表
     */
    public List<String> getSheetInfo(){
        Workbook workbook = workBookContext.getWorkbook();
        return IntStream.range(0, workbook.getNumberOfSheets())
                .mapToObj(workbook::getSheetName)
                .collect(Collectors.toList());
    }

    /**
     * [ROOT]
     * 创建额外扩展
     * 可能会计划出一些扩展功能
     */
    protected void createAdditionalExtensions() {
        // 初始化数据校验器
        try (ValidatorFactory validatorFactory = Validation.buildDefaultValidatorFactory()) {
            this.validator = validatorFactory.getValidator();
        }
    }

    /**
     * [内部方法]
     * 检查文件格式
     */
    protected DetectResult checkFileFormat(File file,InputStream ins){
        // 检查文件格式是否为XLSX
        DetectResult detectResult = this.getFileOrStreamDetectResult(file, ins, TikaShell.OOXML_EXCEL);
        if (!detectResult.isDetect()){
            DetectResult.FileStatus currentFileStatus = detectResult.getCurrentFileStatus();
            if (currentFileStatus == DetectResult.FileStatus.FILE_MIME_TYPE_PROBLEM ||
                    currentFileStatus == DetectResult.FileStatus.FILE_SUFFIX_PROBLEM
            ){
                // 如果是因为后缀不匹配或媒体类型不匹配导致识别不通过 换XLS格式再次识别
                detectResult = this.getFileOrStreamDetectResult(file, ins, TikaShell.MS_EXCEL);
            }else {
                // 不通过抛出异常
                detectResult.throwException();
            }
        }
        // 检查文件是否是需要的类型，否则抛出异常
        if (!detectResult.isDetect()){
            detectResult.throwException();
        }
        return detectResult;
    }

    /**
     * [内部方法]
     *
     * @param file 文件
     * @param ins 流
     * @param mimeType 媒体类型
     * @return 检查结果
     */
    protected DetectResult getFileOrStreamDetectResult(File file, InputStream ins, MimeType mimeType){
        DetectResult detectResult;
        if (file == null){
            detectResult = TikaShell.detect(ins,mimeType,false);
        }else {
            detectResult = TikaShell.detect(file,mimeType,true);
        }
        return detectResult;
    }

    /**
     * [内部方法]
     * 加载文件并读取数据创建Workbook
     */
    protected void loadFileDataToWorkBook() {
        // 读取文件加载到元信息
        try(InputStream fis = new ByteArrayInputStream(workBookContext.getDataCache())){
            Workbook workbook;
            if (workBookContext.getMimeType() == TikaShell.OOXML_EXCEL){
                try {
                    IOUtils.setByteArrayMaxOverride(200000000);
                    this.workBookContext.setEventDriven();
                    OPCPackage opcPackage = OPCPackage.open(fis);
                    workbook = XSSFWorkbookFactory.createWorkbook(opcPackage);
                    opcPackage.close();
                    IOUtils.setByteArrayMaxOverride(-1);
                }catch (IOException e){
                    if (e.getMessage().contains("Zip bomb detected!")){
                        boolean allowLimitProtect = _sheetLevelReaderConfig.getReadPolicyAsBoolean(ExcelReadPolicy.ALLOW_BREAK_THROUGH_RESOURCES_LIMIT_PROTECT);
                        if (allowLimitProtect){
                            ZipSecureFile.setMinInflateRatio(0D);
                            LoggerHelper.warn(LOGGER,"文件大小超出系统限制，已自动开启跳过检测.");
                            this.workBookContext.setEventDriven();
                            OPCPackage opcPackage = OPCPackage.open(fis);
                            workbook = XSSFWorkbookFactory.createWorkbook(opcPackage);
                            opcPackage.close();
                        }else {
                            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,"禁止读取超过限制文件,请检查文件格式");
                        }
                    }else{
                        throw e;
                    }
                }finally {
                    ZipSecureFile.setMinInflateRatio(0.01D);
                }
            }else {
                workbook = WorkbookFactory.create(fis);
            }
            workBookContext.setWorkbook(workbook);
        } catch (IOException | RecordFormatException | InvalidFormatException e) {
            LOGGER.error("加载文件失败",e);
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,e.getMessage());
        }
    }

    /**
     * 读取Excel工作表为一个实体
     * @param readerConfigBuilder 读取配置构建器
     * @param <RT>  读取类型
     * @return 读取的数据
     * @see SpecifyPositionBind 单元格绑定属性
     */
    public <RT> RT readSheetDataAsObject(ReadConfigBuilder<RT> readerConfigBuilder){
        return this.readSheetDataAsObject(readerConfigBuilder.build());
    }

    /**
     * [ROOT]
     * 读取Excel工作表为一个实体
     * @param readerConfig 读取配置
     * @param <RT>  读取类型
     * @return 读取的数据
     * @see SpecifyPositionBind 单元格绑定属性
     */
    public <RT> RT readSheetDataAsObject(ReaderConfig<RT> readerConfig){
        if (readerConfig != null){
            readerConfig.setReadAsObject(true);
        }
        assert readerConfig != null;
        Sheet sheet = this.searchSheet(readerConfig);
        this.preCheckAndFixReadConfig(readerConfig);
        this.spreadMergedCells(sheet,readerConfig);
        RT instance = readerConfig.getCastClassInstance();
        this.convertPositionCellToInstance(instance, readerConfig,sheet);
        this.validateConvertEntity(instance, readerConfig);
        return instance;
    }

    /**
     * @param readerConfig 读取配置
     */
    protected Sheet searchSheet(ReaderConfig<?> readerConfig){
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
     * @see ExcelReadPolicy#SPREAD_MERGING_REGION
     * @param sheet 工作表
     */
    protected void spreadMergedCells(Sheet sheet,ReaderConfig<?> readerConfig) {
        if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.SPREAD_MERGING_REGION)){
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
                        Cell cell = ExcelToolkit.createOrCatchCell(sheet, rowIndex, columnIndex, null);
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

    }

    /**
     * 查找表头
     * @see ColumnBind 列绑定注解
     *
     * @param readerConfig 读取配置
     */
    protected void searchHeaderCellPosition(ReaderConfig<?> readerConfig){
        if (readerConfig.getSheetIndex() == -1){return;}
        Sheet sheet = workBookContext.getIndexSheet(readerConfig.getSheetIndex());
        // 没有表跳过匹配
        if (sheet == null){return;}
        List<EntityCellMappingInfo<?>> indexMappingInfos = readerConfig.getIndexMappingInfos();
        // 提取指定表头
        Map<String, Integer> headerKeys = indexMappingInfos.stream()
                .map(EntityCellMappingInfo::getHeaderName)
                .filter(Objects::nonNull)
                .distinct()
                .collect(Collectors.toMap(element -> element, i -> -1));
        // 实体如果有指定表头则进入匹配表头阶段
        if (!headerKeys.isEmpty()){
            // 从缓存中获取表头映射
            Map<Integer, HashBasedTable<String, Integer, Integer>> headerCaches = workBookContext.getHeaderCaches();
            //                表头名称,序号,列位置
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
            int readHeadRows;
            if (readerConfig.getSearchHeaderMaxRows() > 0){
                readHeadRows = readerConfig.getSearchHeaderMaxRows();
            }else{
                readHeadRows = Math.min(getRecordRowNumber(readerConfig), AxolotlDefaultReaderConfig.XVOID_DEFAULT_HEADER_FINDING_ROW);
            }
            // 拣出已经匹配的表头，如果已经匹配过的表头将不会再匹配
            Map<String, Integer> notAlreadyRecordKeys = new HashMap<>();
            if (hintCache){
                Set<String> alreadyRecordKeySet = headerCache.rowKeySet();
                headerKeys.keySet().stream()
                        .filter(headerKey -> !alreadyRecordKeySet.contains(headerKey))
                        .forEach(headerKey -> notAlreadyRecordKeys.put(headerKey, 0));
            }else {
                notAlreadyRecordKeys.putAll(headerKeys);
            }
            if (!notAlreadyRecordKeys.isEmpty()){
                int[] sheetColumnEffectiveRange = readerConfig.getSheetColumnEffectiveRange();
                for (int i = 0; i < readHeadRows; i++) {
                    Row row = sheet.getRow(i);
                    int endColumnRange = sheetColumnEffectiveRange[1] == -1 ? row.getLastCellNum() : sheetColumnEffectiveRange[1];
                    if (ExcelToolkit.notBlankRowCheck(row, sheetColumnEffectiveRange[0], endColumnRange)){
                        Iterator<Cell> cellIterator = row.cellIterator();
                        while (cellIterator.hasNext()){
                            Cell cell = cellIterator.next();
                            // 应该没有人会用数字当表头吧...
                            if (cell != null && cell.getCellType() == CellType.STRING){
                                // FIX 20240701 移除两边空格
                                String cellValue = cell.getStringCellValue().trim();
                                if (headerKeys.containsKey(cellValue) && notAlreadyRecordKeys.containsKey(cellValue)){
                                    LoggerHelper.debug(LOGGER,LoggerHelper.format("查找到表头[%s]", cellValue));
                                    headerCache.put(cellValue, headerCache.row(cellValue).size()+1 , cell.getColumnIndex());
                                }
                            }
                        }
                    }
                }
            }
            LoggerHelper.debug(LOGGER,LoggerHelper.format("查找表头结束,映射信息:%s", headerCache));
            if (!hintCache){
                headerCaches.put(readerConfig.getSheetIndex(), headerCache);
            }
            // 为映射指定列索引
            for (EntityCellMappingInfo<?> indexMappingInfo : indexMappingInfos) {
                String headerName = indexMappingInfo.getHeaderName();
                if (StringUtils.isNotBlank(headerName)){
                    Map<Integer, Integer> recordInfo = headerCache.row(headerName);
                    if (recordInfo.isEmpty()){
                        if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.IGNORE_EMPTY_SHEET_HEADER_ERROR)){
                            LoggerHelper.debug(LOGGER,LoggerHelper.format("表头[%s]不存在", headerName));
                        }else {
                            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,LoggerHelper.format("表头[%s]不存在", headerName));
                        }
                        continue;
                    }
                    // 从筛选过的key中获取当前批次已经指定过的表头位置
                    Integer assignedIndex = headerKeys.get(headerName);
                    Integer columnIndex;
                    int headerNameIndex = indexMappingInfo.getHeaderNameIndex();
                    // -1没有指定的话按顺序取值
                    if (headerNameIndex == -1){
                        // -1 为第一次匹配取第一个
                        if (assignedIndex == -1){assignedIndex = 1;} else {assignedIndex+=1;}
                        columnIndex = recordInfo.getOrDefault(assignedIndex, -1);
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
    protected <RT> RT readRow(Sheet sheet,int rowNumber,ReaderConfig<RT> readerConfig){
        RT instance = readerConfig.getCastClassInstance();
        Row row = sheet.getRow(rowNumber);
        if (ExcelToolkit.blankRowCheck(row,readerConfig)){
            if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.INCLUDE_EMPTY_ROW)){
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
    protected <RT> void convertCellToInstance(Row row,RT instance,ReaderConfig<RT> readerConfig){
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
    protected <RT> void row2SimplePOJO(RT instance, Row row, ReaderConfig<RT> readerConfig){
        this.convertPositionCellToInstance(instance,readerConfig,row.getSheet());
        List<EntityCellMappingInfo<?>> indexMappingInfos = readerConfig.getIndexMappingInfos();
        for (EntityCellMappingInfo<?> indexMappingInfo : indexMappingInfos) {
            workBookContext.setCurrentReadRowIndex(row.getRowNum());
            workBookContext.setCurrentReadColumnIndex(indexMappingInfo.getColumnPosition());
            CellGetInfo cellValue = this.getCellOriginalValue(row,indexMappingInfo.getColumnPosition(), indexMappingInfo,readerConfig);
            Object adaptiveValue = this.adaptiveCellValue2EntityClass(cellValue, indexMappingInfo, readerConfig);
            this.assignValueToField(instance,adaptiveValue,indexMappingInfo,readerConfig);
        }
        this.validateConvertEntity(instance,readerConfig);
    }

    /**
     * @param readerConfig 读取配置
     */
    protected void convertPositionCellToInstance(Object instance,ReaderConfig<?> readerConfig,Sheet sheet){
        List<EntityCellMappingInfo<?>> positionMappingInfos = readerConfig.getPositionMappingInfos();
        for (EntityCellMappingInfo<?> positionMappingInfo : positionMappingInfos) {
            workBookContext.setCurrentReadRowIndex(positionMappingInfo.getRowPosition());
            workBookContext.setCurrentReadColumnIndex(positionMappingInfo.getColumnPosition());
            CellGetInfo cellGetInfo = this.getPositionCellOriginalValue(sheet, positionMappingInfo);
            Object adaptiveValue = this.adaptiveCellValue2EntityClass(cellGetInfo, positionMappingInfo, readerConfig);
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
    protected void assignValueToField(Object instance,Object adaptiveValue, EntityCellMappingInfo<?> mappingInfo,ReaderConfig<?> readerConfig){
        Field field = ReflectToolkit.recursionGetField(instance.getClass(),mappingInfo.getFieldName());
        assert field != null;
        field.setAccessible(true);
        Object o = field.get(instance);
        boolean useSetter = readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.READ_FIELD_USE_SETTER);
        if (o!= null){
            if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.FIELD_EXIST_OVERRIDE)){
                this.assignOrInvokeSetter(instance, adaptiveValue, useSetter, field);
            }
        }else {
            this.assignOrInvokeSetter(instance, adaptiveValue, useSetter, field);
        }
    }

    /**
     * 根据设定，将给定的值分配给对象的字段或者通过setter方法设置。
     * 如果useSetter为true，并且字段上有AxolotlReaderSetter注解，则尝试通过注解指定的方法名调用相应的方法。
     * 如果没有指定方法名或者useSetter为false，则直接通过反射设置字段值。
     *
     * @param instance 需要设置字段值的对象实例
     * @param adaptiveValue 要分配或设置的值
     * @param useSetter 指定是否使用setter方法进行设置
     * @param field 目标字段
     * @throws IllegalAccessException 当访问字段或调用方法时发生访问权限问题时抛出
     */
    private void assignOrInvokeSetter(Object instance, Object adaptiveValue, boolean useSetter, Field field) throws IllegalAccessException {
        if (useSetter){
            AxolotlReaderSetter readerSetter = field.getAnnotation(AxolotlReaderSetter.class);
            String methodName = null;
            if (readerSetter != null){
                methodName =  readerSetter.value();
            }
            if (methodName != null){
                ReflectToolkit.invokeMethod(methodName, instance, ClassToolkit.castObjectArray2ClassArray(List.of(adaptiveValue)), adaptiveValue);
            }else{
                ReflectToolkit.invokeFieldSetter(field, instance, adaptiveValue);
            }
        }else{
            int modifiers = field.getModifiers();
            if(Modifier.isStatic(modifiers)){
                LoggerHelper.error(LOGGER,"静态字段[%s]不可赋值",field.getName());
                return;
            }
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
    protected Object adaptiveCellValue2EntityClass(CellGetInfo info, EntityCellMappingInfo<?> mappingInfo, ReaderConfig<?> readerConfig){
        if (mappingInfo.getDataCastAdapter() == AutoAdapter.class){
            DataCastAdapter<Object> autoAdapter = AutoAdapter.INSTANCE();
            return this.adaptiveValue(autoAdapter,info, (EntityCellMappingInfo<Object>) mappingInfo, (ReaderConfig<Object>) readerConfig);
        }
        Class<? extends DataCastAdapter<?>> dataCastAdapterClass = mappingInfo.getDataCastAdapter();
        if (dataCastAdapterClass != null && !dataCastAdapterClass.isInterface()){
            DataCastAdapter adapter;
            try {
                Map<Class<?>, DataCastAdapter<?>> castAdapterCache = this.workBookContext.getCastAdapterCache();
                if (castAdapterCache.containsKey(dataCastAdapterClass)){
                    adapter = castAdapterCache.get(dataCastAdapterClass);
                }else {
                    adapter = dataCastAdapterClass.getDeclaredConstructor().newInstance();
                    castAdapterCache.put(dataCastAdapterClass,adapter);
                }
                return this.adaptiveValue(adapter,info, (EntityCellMappingInfo<Object>) mappingInfo, (ReaderConfig<Object>) readerConfig);
            } catch (InstantiationException | IllegalAccessException |
                     InvocationTargetException | NoSuchMethodException e) {
                throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.CONVERT_FIELD_ERROR,e);
            }
        }else {
            throw new AxolotlExcelReadException(mappingInfo,String.format("[%s]字段请配置适配器,字段类型:[%s]",mappingInfo.getFieldName(), mappingInfo.getFieldType()));
        }
    }

    /**
     * @param readerConfig 读取配置
     */
    protected Object adaptiveValue(DataCastAdapter<Object> adapter, CellGetInfo info, EntityCellMappingInfo<Object> mappingInfo, ReaderConfig<Object> readerConfig) {
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
    protected Object castValue(DataCastAdapter<Object> adapter, CellGetInfo info, EntityCellMappingInfo<Object> mappingInfo) {
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
    protected CellGetInfo getPositionCellOriginalValue(Sheet sheet, EntityCellMappingInfo<?> mappingInfo){
        Row row = sheet.getRow(mappingInfo.getRowPosition());
        if (row == null){
            return this.getBlankCellValue(mappingInfo);
        }
        Cell cell = row.getCell(mappingInfo.getColumnPosition());
        if (cell == null){
            return this.getBlankCellValue(mappingInfo);
        }
        return this.getCellOriginalValue(row,mappingInfo.getColumnPosition(), mappingInfo,null);
    }

    /**
     * 获取单元格原始值
     *
     * @param row 行次
     * @param mappingInfo 映射信息
     * @see #getIndexCellValue(Row, int, EntityCellMappingInfo,ReaderConfig)
     * @return 单元格值
     */
    protected CellGetInfo getCellOriginalValue(Row row,int index, EntityCellMappingInfo<?> mappingInfo,ReaderConfig<?> readerConfig){
        // 一般不为null，由map类型传入时，默认使用索引映射
        if (mappingInfo == null){
            mappingInfo = new EntityCellMappingInfo<>(String.class);
            mappingInfo.setColumnPosition(index);
        }
        return this.getIndexCellValue(row,index, mappingInfo,readerConfig);
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
    protected CellGetInfo getIndexCellValue(Row row,int index, EntityCellMappingInfo<?> mappingInfo,ReaderConfig<?> readerConfig){
        // 未找到映射返回空值
        if (index < 0){
            return this.getBlankCellValue(mappingInfo);
        }
        // 不在表有效索引范围中
        if (readerConfig != null){
            int endColumnRange = readerConfig.getSheetColumnEffectiveRange()[1] == -1 ? row.getLastCellNum() : readerConfig.getSheetColumnEffectiveRange()[1];
            if(readerConfig.getSheetColumnEffectiveRange()[0] > index || endColumnRange <= index){
                return this.getBlankCellValue(mappingInfo);
            }
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
    protected CellGetInfo getBlankCellValue(EntityCellMappingInfo<?> mappingInfo){
        CellGetInfo cellGetInfo = new CellGetInfo();
        cellGetInfo.setCellType(CellType.BLANK);
        if (mappingInfo.fieldIsPrimitive()){
            cellGetInfo.setCellValue(mappingInfo.fillDefaultPrimitiveValue(null));
        }
        return cellGetInfo;
    }

    private final static DecimalFormat decimalFormat = new DecimalFormat(StringPool.HASH);

    private static final String MAP_VALUE_PREFIX = "CELL_";
    /**
     * [ROOT]
     * 填充单元格数据到map
     *
     * @param readerConfig 读取配置
     */
    protected <RT> void row2MapInstance(Map<String,Object> instance, Row row,ReaderConfig<RT> readerConfig){
        workBookContext.setCurrentReadRowIndex(row.getRowNum());
        int[] sheetColumnEffectiveRange = readerConfig.getSheetColumnEffectiveRange();
        short endColumnRange = sheetColumnEffectiveRange[1] < 0 ? row.getLastCellNum() : (short) sheetColumnEffectiveRange[1];
        for (int i = sheetColumnEffectiveRange[0]; i < endColumnRange; i++) {
            Cell cell = row.getCell(i);
            if (cell != null){
                workBookContext.setCurrentReadColumnIndex(cell.getColumnIndex());
                int idx = cell.getColumnIndex();

                CellGetInfo cellOriginalValue = getCellOriginalValue(row, cell.getColumnIndex(), null, readerConfig);
                instance.putAll(mapMasterKey(i, cellOriginalValue,readerConfig));

                if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.USE_MAP_DEBUG) && !readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.MAP_CONVERT_INFO_OBJECT)){
                    instance.put("CELL_"+idx+"@TYPE",cell.getCellType());
                }
            }
        }
    }

    /**
     * map万能钥匙
     * @param index 列索引
     * @param cellGetInfo 单元格值
     * @param readerConfig 读取配置
     * @return map读取信息
     * @param <RT> 读取类型
     */
    private <RT> Map<String, Object> mapMasterKey(int index, CellGetInfo cellGetInfo, ReaderConfig<RT> readerConfig) {
        boolean sortedDataPolicy = readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.SORTED_READ_SHEET_DATA);
        boolean mapConvertObjectPolicy = readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.MAP_CONVERT_INFO_OBJECT);
        Map<String, Object> globalInfo = sortedDataPolicy ? new LinkedHashMap<>() : new HashMap<>();
        Map<String, Object> convertedInfo = sortedDataPolicy ? new LinkedHashMap<>() : new HashMap<>();
        String key = MAP_VALUE_PREFIX + index;
        Map<String, MapDocker<?>> mapDockerMap = readerConfig.getMapDockerMap();
        boolean allowPutNullValue = readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.MAP_ALLOW_PUT_NULL_VALUE);
        for (Map.Entry<String, MapDocker<?>> dockerEntry : mapDockerMap.entrySet()) {
            String extendKey = mapConvertObjectPolicy ? dockerEntry.getKey() : key+StringPool.AT+dockerEntry.getKey();
            if (!readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.FIELD_EXIST_OVERRIDE) && convertedInfo.containsKey(extendKey)){
                LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME,String.format("字段:[%s]已存在,跳过",extendKey));
                continue;
            }
            MapDocker<?> docker = dockerEntry.getValue();
            Object convertedValue = docker.convert(index, cellGetInfo, readerConfig);
            if (convertedValue == null){
                Boolean nullDisplay = docker.getNullDisplay();
                if (nullDisplay == null){
                    nullDisplay = allowPutNullValue;
                    LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME,String.format("字段:[%s]为空,是否显示:[使用全局配置]- %s",extendKey,nullDisplay));
                }else{
                    LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME,String.format("字段:[%s]为空,是否显示:%s",extendKey,nullDisplay));
                }
                if (nullDisplay){convertedInfo.put(extendKey,null);}
            }else{
                convertedInfo.put(extendKey, convertedValue);
            }
        }
        if (mapConvertObjectPolicy){
            AxolotlCellMapInfo axolotlCellMapInfo = new AxolotlCellMapInfo(index,cellGetInfo.getCellValue(),cellGetInfo.getCellType());
            if (!convertedInfo.isEmpty()){
                axolotlCellMapInfo.setDockerValues(convertedInfo);
            }
            globalInfo.put(key, axolotlCellMapInfo);
        }else {
            globalInfo.put(key, cellGetInfo.getCellValue());
            globalInfo.putAll(convertedInfo);
        }
        return globalInfo;
    }

    /**
     * 校验读取实体是否符合验证规则
     */
    protected <RT> void validateConvertEntity(RT instance, ReaderConfig<RT> readerConfig) {
        boolean isValidate = readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.VALIDATE_READ_ROW_DATA);
        if (isValidate){
            Set<ConstraintViolation<RT>> validate = validator.validate(instance, readerConfig.getValidGroups());
            if (!validate.isEmpty()){
                for (ConstraintViolation<RT> constraintViolation : validate) {
                    AxolotlExcelReadException axolotlExcelReadException = new AxolotlExcelReadException(workBookContext, constraintViolation.getMessage());
                    axolotlExcelReadException.setExceptionType(AxolotlExcelReadException.ExceptionType.VALIDATION_ERROR);
                    throw axolotlExcelReadException;
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
    public void preCheckAndFixReadConfig(ReaderConfig<?> readerConfig) {
        //检查部分
        if (readerConfig == null){
            String msg = "读取配置不能为空";
            LOGGER.error(msg);
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,msg);
        }
        int sheetIndex = readerConfig.getSheetIndex();
        if (sheetIndex < 0){
            String msg = String.format("读取的sheet不存在[%s]",readerConfig.getSheetName() != null? readerConfig.getSheetName() : readerConfig.getSheetIndex());
            if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.IGNORE_EMPTY_SHEET_ERROR)){
                LoggerToolkitKt.warnWithModule(LOGGER,Meta.MODULE_NAME,msg+"将返回空数据");
                return;
            }
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,msg);
        }
        Sheet indexSheet = workBookContext.getIndexSheet(sheetIndex);
        if (readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.ALLOW_READ_HIDDEN_SHEET)){
            LoggerHelper.warn(LOGGER,"工作表[%s]为隐藏表，请检查数据是否正确",sheetIndex+1);
        }else{
            Workbook workbook = getWorkBookContext().getWorkbook();
            if (workbook.isSheetHidden(sheetIndex) || workbook.isSheetVeryHidden(sheetIndex)){
                String message = LoggerHelper.format("工作表[%s]为隐藏表",sheetIndex+1);
                LoggerHelper.error(LOGGER,message);
                throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,message);
            }
        }
        readerConfig.getReadPolicyAsBoolean(ExcelReadPolicy.ALLOW_READ_HIDDEN_SHEET);
        Class<?> castClass = readerConfig.getCastClass();
        if (castClass == null){
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,"读取的类型对象不能为空");
        }
        if (readerConfig.getStartIndex() < 0){
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,"读取起始行不得小于0");
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
                throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,"读取结束行不得小于0");
            }
            if (!readerConfig.isReadAsObject()){
                LOGGER.info("未设置读取的结束行,将被默认修正为读取该表最大行数");
            }
            readerConfig.setEndIndex(indexSheet.getLastRowNum()+1);
        }
    }

    /**
     * [ROOT]
     * 计算单元格公式为结果
     *
     * @param cell 单元格
     * @return 计算结果
     */
    protected CellGetInfo getFormulaCellValue(Cell cell) {
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
                throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR, msg);
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

}

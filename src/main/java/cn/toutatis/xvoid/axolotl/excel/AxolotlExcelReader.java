package cn.toutatis.xvoid.axolotl.excel;

import cn.toutatis.xvoid.axolotl.excel.constant.AxolotlDefaultConfig;
import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.constant.ReadExcelFeature;
import cn.toutatis.xvoid.axolotl.excel.support.CastContext;
import cn.toutatis.xvoid.axolotl.excel.support.CellGetInfo;
import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.adapters.AbstractDataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.adapters.AutoAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.exceptions.AxolotlReadException;
import cn.toutatis.xvoid.axolotl.excel.support.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.excel.support.tika.TikaShell;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import lombok.Getter;
import lombok.SneakyThrows;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.RecordFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.jetbrains.annotations.NotNull;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;
import java.util.function.Consumer;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class AxolotlExcelReader<T> implements Iterable<List<T>>{

    /**
     * 日志
     */
    private final Logger LOGGER  = LoggerToolkit.getLogger(AxolotlExcelReader.class);

    /**
     * 工作簿元信息
     */
    @Getter
    private WorkBookContext workBookContext;

    /**
     * 构造文件读取器
     */
    public AxolotlExcelReader(File excelFile) {
        this(excelFile,true);
    }

    @SuppressWarnings("unchecked")
    public AxolotlExcelReader(File excelFile, boolean withDefaultConfig) {
        this(excelFile, (Class<T>) Object.class,withDefaultConfig);
    }

    public AxolotlExcelReader(File excelFile, Class<T> clazz) {
        this(excelFile,clazz,true);
    }

    /**
     * [ROOT]
     * 构造文件读取器
     * @param excelFile Excel工作簿文件
     * @param withDefaultConfig 是否使用默认配置
     */
    public AxolotlExcelReader(File excelFile, Class<T> clazz, boolean withDefaultConfig) {
        if (clazz == null){
            throw new IllegalArgumentException("读取的类型对象不能为空");
        }
        this.detectFileAndInitWorkbook(excelFile);
        this.workBookContext.setDirectReadClass(clazz);
        this.workBookContext.setUseDefaultReaderConfig(withDefaultConfig);
    }

    /**
     * 初始化读取Excel文件
     * 1.初始化加载文件先判断文件是否正常并且是需要的格式
     * 2.将文件加载到POI工作簿中
     * @param excelFile Excel工作簿文件
     */
    private void detectFileAndInitWorkbook(File excelFile) {
        // 检查文件格式是否为XLSX
        DetectResult detectResult = TikaShell.detect(excelFile, TikaShell.OOXML_EXCEL,false);
        if (!detectResult.isDetect()){
            DetectResult.FileStatus currentFileStatus = detectResult.getCurrentFileStatus();
            if (currentFileStatus == DetectResult.FileStatus.FILE_MIME_TYPE_PROBLEM ||
                    currentFileStatus == DetectResult.FileStatus.FILE_SUFFIX_PROBLEM
            ){
                //如果是因为后缀不匹配或媒体类型不匹配导致识别不通过 换XLS格式再次识别
                detectResult = TikaShell.detect(excelFile, TikaShell.MS_EXCEL,false);
            }else {
                //如果是预检查不通过抛出异常
                detectResult.throwException();
            }
        }
        // 检查文件是否正常并且是需要的类型，否则抛出异常
        if (detectResult.isDetect()){
            workBookContext = new WorkBookContext(excelFile,detectResult);
        }else{
            detectResult.throwException();
        }
        // 读取文件加载到元信息
        try(FileInputStream fis = new FileInputStream(workBookContext.getFile())){
            // 校验文件大小
            Workbook workbook;
            if (detectResult.getCatchMimeType() == TikaShell.OOXML_EXCEL && excelFile.length() > AxolotlDefaultConfig.XVOID_DEFAULT_MAX_FILE_SIZE){
                this.workBookContext.setEventDriven(true);
                OPCPackage opcPackage = OPCPackage.open(fis);
                workbook = XSSFWorkbookFactory.createWorkbook(opcPackage);
                opcPackage.close();
            }else {
                workbook = WorkbookFactory.create(fis);
            }
            workBookContext.setWorkbook(workbook);
        } catch (IOException | RecordFormatException | InvalidFormatException e) {
            LOGGER.error("加载文件失败",e);
            throw new AxolotlReadException(e.getMessage());
        }
    }

    /**
     * [ROOT]
     * 读取Excel数据
     * @param readerConfig 读取配置
     * @return 读取数据
     * @param <RT> 读取的类型泛型
     */
    public <RT> List<RT> readSheetData(ReaderConfig<RT> readerConfig) {
        // 检查并修正配置文件
        this.preCheckAndFixReadConfig(readerConfig);
        List<RT> readResult = new ArrayList<>();
        Sheet sheet = workBookContext.getWorkbook().getSheetAt(readerConfig.getSheetIndex());
        if (sheet == null){
            String msg = "读取的sheet不存在[%s]".formatted(readerConfig.getSheetIndex());
            if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.IGNORE_EMPTY_SHEET_ERROR)){
                LOGGER.warn(msg+"将返回空数据");
                return readResult;
            }else{
                LOGGER.error(msg);
                throw new AxolotlReadException(msg);
            }
        }
        this.readSheetData(sheet,readerConfig,readResult);
        return readResult;
    }

    /**
     * 读取表中每一行的数据
     */
    private <RT> void readSheetData(Sheet sheet,ReaderConfig<RT> readerConfig,List<RT> list){
        int startIndex = readerConfig.getStartIndex();
        int endIndex = readerConfig.getEndIndex();
        if (startIndex == 0){
            int initialRowPositionOffset = readerConfig.getInitialRowPositionOffset();
            if (initialRowPositionOffset > 0){
                LOGGER.debug("跳过前{}行",initialRowPositionOffset);
                startIndex = startIndex + initialRowPositionOffset;
                endIndex = endIndex + initialRowPositionOffset;
            }
        }
        for (int i = startIndex; i < endIndex; i++) {
            RT instance = this.readRow(sheet, i, readerConfig);
            if (instance!= null){list.add(instance);}
        }
    }

    /**
     * [ROOT]
     * 读取行信息到对象
     * @param sheet 表
     * @param rowNumber 行号
     * @param readerConfig 读取配置
     * @param <RT> 转换类型
     */
    private <RT> RT readRow(Sheet sheet,int rowNumber,ReaderConfig<RT> readerConfig){
        RT instance = readerConfig.getCastClassInstance();
        Row row = sheet.getRow(rowNumber);
        if (row == null){
            if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.INCLUDE_EMPTY_ROW)){
                return instance;
            }else{
                return null;
            }
        }
        this.convertCellToInstance(row,instance,readerConfig);
        return instance;
    }

    @SuppressWarnings({"unchecked","rawtypes"})
    private <RT> void convertCellToInstance(Row row,RT instance,ReaderConfig<RT> readerConfig){
        if (instance instanceof Map mapInstance){
            this.row2MapInstance(mapInstance, row,readerConfig);
        }else{
            this.row2SimplePOJO(instance,row,readerConfig);
        }
    }

    /**
     * 填充单元格数据到Java对象
     * @param instance 对象实例
     * @param row 行
     * @param readerConfig 读取配置
     * @param <RT> 读取类型
     */
    @SneakyThrows
    private <RT> void row2SimplePOJO(RT instance, Row row, ReaderConfig<RT> readerConfig){
        Map<String, EntityCellMappingInfo<?>> positionMappingInfos = readerConfig.getPositionMappingInfos();
        for (Map.Entry<String, EntityCellMappingInfo<?>> positionMappingInfoEntry : positionMappingInfos.entrySet()) {
            EntityCellMappingInfo<?> positionMappingInfo = positionMappingInfoEntry.getValue();
            workBookContext.setCurrentReadRowIndex(positionMappingInfo.getRowPosition());
            workBookContext.setCurrentReadColumnIndex(positionMappingInfo.getColumnPosition());
            CellGetInfo cellValue = this.getPositionCellOriginalValue(row.getSheet(), positionMappingInfo);
            Object adaptiveValue = this.adaptiveCellValue2EntityClass(cellValue, positionMappingInfo, readerConfig);
            Field field = instance.getClass().getDeclaredField(positionMappingInfo.getFieldName());
            field.setAccessible(true);
            field.set(instance, adaptiveValue);
        }
        List<EntityCellMappingInfo<?>> indexMappingInfos = readerConfig.getIndexMappingInfos();
        for (EntityCellMappingInfo<?> indexMappingInfo : indexMappingInfos) {
            workBookContext.setCurrentReadRowIndex(row.getRowNum());
            workBookContext.setCurrentReadColumnIndex(indexMappingInfo.getColumnPosition());
            CellGetInfo cellValue = this.getCellOriginalValue(row,indexMappingInfo.getColumnPosition(), indexMappingInfo);
            Object adaptiveValue = this.adaptiveCellValue2EntityClass(cellValue, indexMappingInfo, readerConfig);
            Field field = instance.getClass().getDeclaredField(indexMappingInfo.getFieldName());
            field.setAccessible(true);
            Object o = field.get(instance);
            if (o!= null){
                if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.FIELD_EXIST_OVERRIDE)){
                    field.set(instance, adaptiveValue);
                }
            }else {
                field.set(instance, adaptiveValue);
            }
        }
    }

    /**
     * 适配实体类的字段
     *
     * @param info 单元格值
     * @param mappingInfo 映射信息
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
                throw new AxolotlReadException(e);
            }
        }else {
            throw new AxolotlReadException("[%s]字段请配置适配器,字段类型:[%s]".formatted(mappingInfo.getFieldName(), mappingInfo.getFieldType()));
        }
    }

    private Object adaptiveValue(DataCastAdapter<Object> adapter, CellGetInfo info, EntityCellMappingInfo<Object> mappingInfo, ReaderConfig<Object> readerConfig) {
        if (adapter == null){
            throw new AxolotlReadException("未找到转换的类型:[%s->%s],字段:[%s]".formatted(info.getCellType(), mappingInfo.getFieldType(), mappingInfo.getFieldName()));
        }
        if (adapter instanceof AbstractDataCastAdapter<Object> abstractDataCastAdapter){
            abstractDataCastAdapter.setReaderConfig(readerConfig);
            abstractDataCastAdapter.setEntityCellMappingInfo(mappingInfo);
            return castValue(abstractDataCastAdapter, info, mappingInfo);
        }
        return castValue(adapter, info, mappingInfo);
    }

    private Object castValue(DataCastAdapter<Object> adapter, CellGetInfo info, EntityCellMappingInfo<Object> mappingInfo) {
        if (!adapter.support(info.getCellType(), mappingInfo.getFieldType())){
            throw new AxolotlReadException("不支持转换的类型:[%s->%s],字段:[%s]".formatted(info.getCellType(), mappingInfo.getFieldType(), mappingInfo.getFieldName()));
        }
        CastContext<Object> castContext = new CastContext<>(mappingInfo.getFieldType(), mappingInfo.getFormat());
        return adapter.cast(info, castContext);
    }


    /**
     * 获取位置映射单元格原始值
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
     * @param row 行次
     * @param mappingInfo 映射信息
     * @return 单元格值
     * @see #getIndexCellValue(Row, int, EntityCellMappingInfo)
     */
    private CellGetInfo getCellOriginalValue(Row row,int index, EntityCellMappingInfo<?> mappingInfo){
        // 一般不为null，由map类型传入时，默认使用索引映射
        if (mappingInfo == null){
            mappingInfo = new EntityCellMappingInfo<>(String.class);
            mappingInfo.setColumnPosition(index);
        }
        if (mappingInfo.getMappingType() == EntityCellMappingInfo.MappingType.POSITION){
            throw new AxolotlReadException("位置映射请使用getPositionCellValue方法");
        }
        return this.getIndexCellValue(row,index, mappingInfo);
    }

    /**
     * 获取索引映射单元格值
     *
     * @param row 行次
     * @param mappingInfo 映射信息
     * @return 单元格值
     * @see #getBlankCellValue(EntityCellMappingInfo)
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
            case STRING -> value = cell.getStringCellValue();
            case NUMERIC -> value = cell.getNumericCellValue();
            case BOOLEAN -> value = cell.getBooleanCellValue();
            case FORMULA -> value = getFormulaCellValue(cell);
            default -> LOGGER.error("未知的单元格类型:{},{}",cell.getCellType(), cell);
        }
        cellGetInfo.setUseCellValue(true);
        cellGetInfo.setCellValue(value);
        return cellGetInfo;
    }

    /**
     * 获取空单元格值
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
     */
    private <RT> void row2MapInstance(Map<String,Object> instance, Row row,ReaderConfig<RT> readerConfig){
        workBookContext.setCurrentReadRowIndex(row.getRowNum());
        row.cellIterator().forEachRemaining(cell -> {
            workBookContext.setCurrentReadColumnIndex(cell.getColumnIndex());
            int idx = cell.getColumnIndex() + 1;
            String key = "CELL_" + idx;
            instance.put(key, getCellOriginalValue(row, cell.getColumnIndex(),null).getCellValue());
            if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.USE_MAP_DEBUG)){
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
     * [ROOT]
     * 预校验读取配置是否正常
     * 不正常的数据将被修正
     * @param readerConfig 读取配置
     */
    private void preCheckAndFixReadConfig(ReaderConfig<?> readerConfig) {
        //检查部分
        if (readerConfig == null){
            String msg = "读取配置不能为空";
            LOGGER.error(msg);
            throw new AxolotlReadException(msg);
        }
        int sheetIndex = readerConfig.getSheetIndex();
        if (sheetIndex < 0){
            throw new AxolotlReadException("读取的sheetIndex不能小于0");
        }
        Class<?> castClass = readerConfig.getCastClass();
        if (castClass == null){
            throw new AxolotlReadException("读取的类型对象不能为空");
        }
        if (readerConfig.getStartIndex() < 0){
            throw new AxolotlReadException("读取起始行不得小于0");
        }
        if (readerConfig.getEndIndex() < 0){
            throw new AxolotlReadException("读取结束行不得小于0");
        }
        //修正部分
        if (readerConfig.getInitialRowPositionOffset() < 0){
            LOGGER.warn("读取的初始行偏移量不能小于0，将被修正为0");
            readerConfig.setInitialRowPositionOffset(0);
        }

    }

    /**
     * [ROOT]
     * 计算单元格公式为结果
     * @param cell 单元格
     * @return 计算结果
     */
    private CellGetInfo getFormulaCellValue(Cell cell) {
        // 从元数据中获取计算计算器
        CellValue evaluated = workBookContext.getFormulaEvaluator().evaluate(cell);
        // 将单元格为公式的单元格值转换为计算结果
        Object value = switch (evaluated.getCellType()) {
            case STRING -> evaluated.getStringValue();
            case NUMERIC -> evaluated.getNumberValue();
            case BOOLEAN -> evaluated.getBooleanValue();
            default -> {
                String msg = String.format("未知的公式单元格类型位置:[%d,%d],单元格类型:[%s],单元格值:[%s]",
                        cell.getRowIndex(), cell.getColumnIndex(), evaluated.getCellType(), evaluated);
                LOGGER.error(msg);
                throw new AxolotlReadException(msg);
            }
        };
        CellGetInfo cellGetInfo = new CellGetInfo(true, value);
        cellGetInfo.setCellType(cell.getCellType());
        return cellGetInfo;
    }

    @NotNull
    @Override
    public Iterator<List<T>> iterator() {
        // TODO 迭代器
        return null;
    }

    @Override
    public void forEach(Consumer<? super List<T>> action) {
        Iterable.super.forEach(action);
    }

    @Override
    public Spliterator<List<T>> spliterator() {
        return Iterable.super.spliterator();
    }
}

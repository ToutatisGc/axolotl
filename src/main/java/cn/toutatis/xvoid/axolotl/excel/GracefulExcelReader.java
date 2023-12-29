package cn.toutatis.xvoid.axolotl.excel;

import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.constant.ReadExcelFeature;
import cn.toutatis.xvoid.axolotl.excel.support.*;
import cn.toutatis.xvoid.axolotl.excel.support.adapters.AbstractDataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.adapters.DefaultAdapters;
import cn.toutatis.xvoid.axolotl.excel.support.tika.DetectResult;
import cn.toutatis.xvoid.axolotl.excel.support.tika.TikaShell;
import cn.toutatis.xvoid.toolkit.constant.Time;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Getter;
import lombok.SneakyThrows;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class GracefulExcelReader<T> {

    /**
     * 日志
     */
    private final Logger LOGGER  = LoggerToolkit.getLogger(GracefulExcelReader.class);

    /**
     * 工作簿元信息
     */
    @Getter
    private WorkBookMetaInfo workBookMetaInfo;

    /**
     * 读取配置集合
     */
    private Map<Class<?>,ReaderConfig<?>> readerConfigMap = new HashMap<>();

    /**
     * 构造文件读取器
     */
    public GracefulExcelReader(File excelFile) {
        this(excelFile,true);
    }


    /**
     * 构造文件读取器
     *
     * @param excelFile Excel工作簿文件
     * @param withDefaultConfig 是否使用默认配置
     */
    public GracefulExcelReader(File excelFile,boolean withDefaultConfig) {
        this.initWorkbook(excelFile);
        this.workBookMetaInfo.setUseDefaultReaderConfig(withDefaultConfig);
    }

    public GracefulExcelReader(File excelFile, Class<T> clazz,boolean withDefaultConfig) {
        this.initWorkbook(excelFile);
        if (clazz == null){
            throw new IllegalArgumentException("读取的类型对象不能为空");
        }
        this.workBookMetaInfo.setDirectClass(clazz);
        this.workBookMetaInfo.setUseDefaultReaderConfig(withDefaultConfig);
    }


    /**
     * [ROOT] 读取表数据
     * @param config 读取配置
     * @return 表数据
     * @param <RT> 读取类型
     */
    public <RT> List<RT> readSheetData(ReaderConfig<RT> config) {
        ReaderConfig<RT> readerConfig = config;
        if (readerConfig == null){
            throw new IllegalArgumentException("读取配置不能为空");
        }

    }

    /**
     * @param sheetIndex 表索引
     */
    @SuppressWarnings("unchecked")
    public <RT> List<T> readSheetData(int sheetIndex,Class<T> clazz,int start ,int end) {
        if (clazz == null || clazz == Object.class){
            throw new IllegalArgumentException("读取的类型对象不能为空");
        }
        // sheetIndex < 0 代表该文件中不存在该sheet
        if (sheetIndex < 0){
            if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.IGNORE_EMPTY_SHEET_ERROR)){
                return null;
            }else{
                String msg = readerConfig.getSheetName() != null ? "表名[" + readerConfig.getSheetName() + "]不存在" : "表索引[" + sheetIndex + "]不存在";
                LOGGER.error(msg);
                throw new IllegalArgumentException(msg);
            }
        }
        readerConfig.setCastClass(clazz);
        readerConfig.setSheetIndex(sheetIndex);
        Sheet sheetAt = workBookMetaInfo.getWorkbook().getSheetAt(sheetIndex);
        // TODO 分页加载数据
        int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
        int lastRowNum = sheetAt.getLastRowNum();
        return (List<T>) loadData(0,lastRowNum);
    }

    /**
     * @param sheetName 表名
     */
    public List<T> readSheetData(String sheetName, Class<T> clazz,int start ,int end) {
        if (Validator.strIsBlank(sheetName)){throw new IllegalArgumentException("表名不能为空");}
        readerConfig.setSheetName(sheetName);
        int sheetIndex = this.workBookMetaInfo.getWorkbook().getSheetIndex(sheetName);
        return readSheetData(sheetIndex,clazz,start,end);
    }

    /**
     *
     */
    public void readClassAsList(){
        //TODO 直接根据class获取信息
//        IndexWorkSheet declaredAnnotation = this.castClass.getDeclaredAnnotation(IndexWorkSheet.class);
    }

    /**
     * 初始化读取Excel文件
     * 1.初始化加载文件先判断文件是否正常并且是需要的格式
     * 2.将文件加载到POI工作簿中
     * @param excelFile Excel工作簿文件
     */
    private void initWorkbook(File excelFile) {
        // 检查文件是否正常
        TikaShell.preCheckFileNormalThrowException(excelFile);
        DetectResult detectResult = TikaShell.detect(excelFile, TikaShell.OOXML_EXCEL,true);
        if (!detectResult.isDetect()){
            // 没有识别到XLSX格式再尝试识别XLS格式
            DetectResult.FileStatus currentFileStatus = detectResult.getCurrentFileStatus();
            if (currentFileStatus == DetectResult.FileStatus.FILE_MIME_TYPE_PROBLEM ||
                    currentFileStatus == DetectResult.FileStatus.FILE_SUFFIX_PROBLEM
            ){
                detectResult = TikaShell.detect(excelFile, TikaShell.MS_EXCEL,true);
            }else {
                detectResult.throwException();
            }
        }
        // 检查文件是否正常并且是需要的类型，否则抛出异常
        if (detectResult.isDetect() && detectResult.isWantedMimeType()){
            workBookMetaInfo = new WorkBookMetaInfo(excelFile,detectResult);
        }else{
            detectResult.throwException();
        }
        // 读取文件加载到元信息
        try(FileInputStream fis = new FileInputStream(workBookMetaInfo.getFile())){
            Workbook workbook = WorkbookFactory.create(fis);
            workBookMetaInfo.setWorkbook(workbook);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 加载数据
     *
     * @param start 开始行
     * @param end 结束行
     */
    private List<Object> loadData(int start,int end){
        // 读取指定sheet
        Sheet sheet = workBookMetaInfo.getWorkbook().getSheetAt(readerConfig.getSheetIndex());
        int sheetIndex = readerConfig.getSheetIndex();
        List<Row> rowList;
        // 缓存sheet数据
        if (workBookMetaInfo.isSheetDataEmpty(sheetIndex)){
            ArrayList<Row> tmp = new ArrayList<>();
            // 是否读取所有行
            if (readerConfig.getReadFeatureAsBoolean(ReadExcelFeature.INCLUDE_EMPTY_ROW)){
                for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    LOGGER.debug(i+":"+row);
                    tmp.add(row);
                }
            }else{
                sheet.rowIterator().forEachRemaining(tmp::add);
            }
            workBookMetaInfo.setSheetData(sheetIndex,tmp);
            rowList = tmp;
        }else {
            rowList = workBookMetaInfo.getSheetData(sheetIndex);
        }
        // 读取分页数据
        List<Object> dataList = new ArrayList<>();
        for (int idx = start; idx <= end; idx++) {
            Object castClassInstance = readerConfig.getCastClassInstance();
            Row row = rowList.get(idx);
            if (row == null){
                dataList.add(castClassInstance);
                continue;
            }
            // 填充对象数据
            putRowToInstance(castClassInstance,row);
            dataList.add(castClassInstance);
        }
        return dataList;
    }

    /**
     * 填充单元格数据到对象
     *
     * @param instance 实例对象
     * @param row 当前行
     */
    @SneakyThrows
    @SuppressWarnings({"unchecked"})
    private void putRowToInstance(Object instance, Row row){
        // 填充到map
        if (instance instanceof Map<?,?> info){
            this.putRowToMapInstance((Map<String, Object>) info,row);
        }else{
            Class<?> castClass = readerConfig.getCastClass();
            Map<String, EntityCellMappingInfo<?>> positionMappingInfos = readerConfig.getPositionMappingInfos();
            List<EntityCellMappingInfo<?>> indexMappingInfos = readerConfig.getIndexMappingInfos();
            for (EntityCellMappingInfo<?> mappingInfo : indexMappingInfos) {
                Field field = castClass.getDeclaredField(mappingInfo.getFieldName());
                field.setAccessible(true);
                // 1. 获取单元格值
                int columnPosition = mappingInfo.getColumnPosition();
                CellGetInfo cellValue;
                if (columnPosition == -1){
                    cellValue = new CellGetInfo(false,mappingInfo.fillDefaultPrimitiveValue(null));
                }else{
                    cellValue = getCellValue(row.getCell(columnPosition), mappingInfo);
                }
                LOGGER.debug(mappingInfo.getFieldName() + ":" + cellValue);
                // 2. 转换单元格值
                Object adaptiveValue = this.adaptiveEntityClass(cellValue, mappingInfo);
                LOGGER.debug("转换前："+cellValue.getCellValue()+" 转换后："+adaptiveValue);
                // 3. 设置单元格值到实体
//                field.set(instance,adaptiveValue);
            }
        }
    }

    /**
     * 适配实体类的字段
     *
     * @param info 单元格值
     * @param mappingInfo 映射信息
     * @param <FT>    实体类
     * @return 适配实体类的字段值
     */
    @SuppressWarnings("unchecked")
    private <FT> Object adaptiveEntityClass(CellGetInfo info, EntityCellMappingInfo<FT> mappingInfo){
        Class<? extends DataCastAdapter<?>> dataCastAdapter = mappingInfo.getDataCastAdapter();
        DataCastAdapter<FT> adapter;
        if (dataCastAdapter != null && !dataCastAdapter.isInterface()){
            try {
                adapter =(DataCastAdapter<FT>) dataCastAdapter.getDeclaredConstructor().newInstance();
            } catch (InstantiationException | IllegalAccessException |
                     InvocationTargetException | NoSuchMethodException e) {
                throw new RuntimeException(e);
            }
        }else {
            DataCastAdapter<?> tmpAdapter = DefaultAdapters.getAdapter(mappingInfo.getFieldType());
            LOGGER.debug("查找默认的适配器类型：{},Class:{}",mappingInfo.getFieldType(),tmpAdapter == null ? "无" : tmpAdapter.getClass().getName());
            if (tmpAdapter == null){
                if(info.getCellType() == null){
                    return info.getCellValue();
                }
                throw new RuntimeException("未找到转换的类型:["+info.getCellType() +"->"+mappingInfo.getFieldType() +" 字段:["+ mappingInfo.getFieldName() +"]");
            }
            AbstractDataCastAdapter<T,FT> abstractDataCastAdapter = (AbstractDataCastAdapter<T,FT>) tmpAdapter;
            abstractDataCastAdapter.setReaderConfig(readerConfig);
            adapter = abstractDataCastAdapter;
        }
        CastContext<FT> castContext =  new CastContext<>(mappingInfo.getFieldType(),mappingInfo.getFormat());
        if (adapter.support(info.getCellType(), mappingInfo.getFieldType())){
            return adapter.cast(info, castContext);
        }else {
            throw new RuntimeException("不支持转换的类型:["+info.getCellType() +"->"+mappingInfo.getFieldType() +" 字段:["+ mappingInfo.getFieldName() +"]");
        }
    }

    /**
     * 填充单元格数据到map
     */
    private void putRowToMapInstance(Map<String,Object> instance, Row row){
        row.cellIterator().forEachRemaining(cell -> {
            int idx = cell.getColumnIndex() + 1;
            String key = "CELL_" + idx;
            instance.put(key,getCellValue(cell, null).getCellValue());
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
     * 获取单元格值
     *
     * @param cell 单元格
     * @param mappingInfo 映射信息
     * @return 单元格值
     */
    private CellGetInfo getCellValue(Cell cell, EntityCellMappingInfo<?> mappingInfo){
        // 一般不为null，又map类型传入时，默认使用索引映射
        if (mappingInfo == null){
            mappingInfo = new EntityCellMappingInfo<>(String.class);
            mappingInfo.setColumnPosition(cell.getColumnIndex());
        }
        return switch (mappingInfo.getMappingType()) {
            case INDEX,UNKNOWN -> this.getIndexCellValue(cell, mappingInfo);
            case POSITION -> null;
            // TODO 位置读取
        };
    }

    private CellGetInfo getPositionCellValue(EntityCellMappingInfo<?> mappingInfo){
        Sheet sheet = workBookMetaInfo.getWorkbook().getSheetAt(readerConfig.getSheetIndex());
        int rowPosition = mappingInfo.getRowPosition();
        if (rowPosition == -1){
            throw new RuntimeException("未设置行位置");
        }
        Row row = sheet.getRow(rowPosition);
        if(row == null){
            // TODO 读取行位置
        }
        return new CellGetInfo();
    }

    /**
     * 获取索引映射单元格值
     *
     * @param cell 单元格
     * @param mappingInfo 映射信息
     * @return 单元格值
     */
    private CellGetInfo getIndexCellValue(Cell cell, EntityCellMappingInfo<?> mappingInfo){
        //TODO 正确读取单元格值
        // 未设置列位置返回空值
        if (mappingInfo.getColumnPosition() == -1){
            // 字段是否为基本类型,基本类型返回默认值
            if (mappingInfo.fieldIsPrimitive()){
                CellGetInfo cellGetInfo = new CellGetInfo();
                cellGetInfo.setUseCellValue(true);
                cellGetInfo.setCellValue(mappingInfo.fillDefaultPrimitiveValue(null));
                return cellGetInfo;
            }
            return new CellGetInfo();
        }
        if (cell == null){
            if (mappingInfo.fieldIsPrimitive()){
                CellGetInfo cellGetInfo = new CellGetInfo();
                cellGetInfo.setUseCellValue(true);
                cellGetInfo.setCellValue(mappingInfo.fillDefaultPrimitiveValue(null));
                return cellGetInfo;
            }else return new CellGetInfo();
        }else {
            Object value = null;
            CellGetInfo cellGetInfo = new CellGetInfo();
            CellType cellType = cell.getCellType();
            cellGetInfo.setCellType(cellType);
            switch (cellType) {
                case STRING -> value = cell.getStringCellValue();
                case NUMERIC -> value = cell.getNumericCellValue();
                case BOOLEAN -> value = cell.getBooleanCellValue();
                case FORMULA -> value = getFormulaCellValue(cell);
                default -> {
                    LOGGER.error("未知的单元格类型:{},{}",cell.getCellType(), cell);
                }
            };
            cellGetInfo.setUseCellValue(true);
            cellGetInfo.setCellValue(value);
            return cellGetInfo;
        }
    }

    /**
     * 计算公式
     * @param cell 单元格
     * @return 计算结果
     */
    private Object getFormulaCellValue(Cell cell){
        CellValue evaluated = workBookMetaInfo.getFormulaEvaluator().evaluate(cell);
        return switch (evaluated.getCellType()) {
            case STRING -> evaluated.getStringValue();
            case NUMERIC -> evaluated.getNumberValue();
            case BOOLEAN -> evaluated.getBooleanValue();
            default -> {
                LOGGER.error("未知的单元格类型:{},{}",evaluated.getCellType(), evaluated);
                yield null;
            }
        };
    }

    public List<T> readSheetData(int sheetIndex, int start, int end) {
        // TODO 读取指定sheet的数据
        return new ArrayList<>();
    }
}

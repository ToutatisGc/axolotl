package cn.toutatis.xvoid.axolotl;

import cn.toutatis.xvoid.axolotl.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.constant.ReadExcelFeature;
import cn.toutatis.xvoid.axolotl.support.*;
import cn.toutatis.xvoid.axolotl.support.adapters.AbstractDataCastAdapter;
import cn.toutatis.xvoid.axolotl.support.adapters.DefaultAdapters;
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
import java.util.List;
import java.util.Map;

/**
 * Excel读取器
 * @author Toutatis_Gc
 */
public class GracefulExcelReader {

    /**
     *
     */
    private final Logger LOGGER  = LoggerToolkit.getLogger(GracefulExcelReader.class);

    /**
     *
     */
    @Getter
    private WorkBookMetaInfo workBookMetaInfo;

    /**
     *
     */
    @Getter
    @SuppressWarnings("rawtypes")
    private final WorkBookReaderConfig workBookReaderConfig;

    /**
     *
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
        this(excelFile,withDefaultConfig,0);
    }

    /**
     * 构造文件读取器
     *
     * @param excelFile Excel工作簿文件
     * @param withDefaultConfig 是否使用默认配置
     * @param initialRowPositionOffset 初始行偏移量
     */
    public GracefulExcelReader(File excelFile,boolean withDefaultConfig,int initialRowPositionOffset) {
        this.initWorkbook(excelFile);
        workBookReaderConfig = new WorkBookReaderConfig<>(withDefaultConfig);
        workBookReaderConfig.setInitialRowPositionOffset(initialRowPositionOffset);
    }

    /**
     * @param sheetIndex 表索引
     */
    @SuppressWarnings("unchecked")
    public <T> List<T> readSheetData(int sheetIndex,Class<T> clazz) {
        if (clazz == null || clazz == Object.class){
            throw new IllegalArgumentException("读取的类型对象不能为空");
        }
        if (sheetIndex < 0){
            boolean ignoreEmptySheetError = workBookReaderConfig.getReadFeatureAsBoolean(ReadExcelFeature.IGNORE_EMPTY_SHEET_ERROR);
            if (ignoreEmptySheetError){
                return null;
            }else{
                String msg = workBookReaderConfig.getSheetName() != null ? "表名[" + workBookReaderConfig.getSheetName() + "]不存在" : "表索引[" + sheetIndex + "]不存在";
                LOGGER.error(msg);
                throw new IllegalArgumentException(msg);
            }
        }
        workBookReaderConfig.setCastClass(clazz);
        workBookReaderConfig.setSheetIndex(sheetIndex);
        Sheet sheetAt = workBookMetaInfo.getWorkbook().getSheetAt(sheetIndex);
        int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
        int lastRowNum = sheetAt.getLastRowNum();
        return (List<T>) loadData(0,lastRowNum);
    }

    /**
     * @param sheetName 表名
     */
    public <T> List<T> readSheetData(String sheetName, Class<T> clazz) {
        if (Validator.strIsBlank(sheetName)){throw new IllegalArgumentException("表名不能为空");}
        workBookReaderConfig.setSheetName(sheetName);
        int sheetIndex = this.workBookMetaInfo.getWorkbook().getSheetIndex(sheetName);
        return readSheetData(sheetIndex,clazz);
    }

    /**
     *
     */
    public void readClass(){
        //TODO 直接根据class获取信息
//        IndexWorkSheet declaredAnnotation = this.castClass.getDeclaredAnnotation(IndexWorkSheet.class);
    }

    /**
     * 读取Excel文件
     *
     * @param excelFile Excel工作簿文件
     */
    private void initWorkbook(File excelFile) {
        // 检查文件是否正常
        TikaShell.preCheckFileNormalThrowException(excelFile);
        DetectResult detectResult = TikaShell.detect(excelFile, TikaShell.OOXML_EXCEL,true);
        if (!detectResult.isDetect()){
            if (detectResult.getCurrentFileStatus() == DetectResult.FileStatus.FILE_MIME_TYPE_PROBLEM ||
                    detectResult.getCurrentFileStatus() == DetectResult.FileStatus.FILE_SUFFIX_PROBLEM
            ){
                detectResult = TikaShell.detect(excelFile, TikaShell.MS_EXCEL,true);
            }else {
                detectResult.throwException();
            }
        }
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
        Sheet sheet = workBookMetaInfo.getWorkbook().getSheetAt(workBookReaderConfig.getSheetIndex());
        int sheetIndex = workBookReaderConfig.getSheetIndex();
        List<Row> rowList;
        // 缓存sheet数据
        if (workBookMetaInfo.isSheetDataEmpty(sheetIndex)){
            ArrayList<Row> tmp = new ArrayList<>();
            // 是否读取所有行
            if (workBookReaderConfig.getReadFeatureAsBoolean(ReadExcelFeature.INCLUDE_EMPTY_ROW)){
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
            Object castClassInstance = workBookReaderConfig.getCastClassInstance();
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
    @SuppressWarnings({"unchecked","rawtypes"})
    private void putRowToInstance(Object instance, Row row){
        // 填充到map
        if (instance instanceof Map info){
            this.putRowToMapInstance(info,row);
        }else{
            Class castClass = workBookReaderConfig.getCastClass();
            List<EntityCellMappingInfo> mappingInfos = workBookReaderConfig.getIndexMappingInfos();
            for (EntityCellMappingInfo mappingInfo : mappingInfos) {
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
                System.err.println(this.adaptiveEntityClass(cellValue,mappingInfo));
                // 3. 设置单元格值到实体
            }

        }
    }

    /**
     * 适配实体类的字段
     *
     * @param info 单元格值
     * @param mappingInfo 映射信息
     * @param <T>  实体类
     * @return 适配实体类的字段值
     */
    private <T> Object adaptiveEntityClass(CellGetInfo info, EntityCellMappingInfo<T> mappingInfo){
        Class<? extends DataCastAdapter<?>> dataCastAdapter = mappingInfo.getDataCastAdapter();
        DataCastAdapter<?> adapter;
        if (dataCastAdapter != null && !dataCastAdapter.isInterface()){
            try {
                adapter = dataCastAdapter.getDeclaredConstructor().newInstance();
            } catch (InstantiationException | IllegalAccessException |
                     InvocationTargetException | NoSuchMethodException e) {
                throw new RuntimeException(e);
            }
        }else {
            AbstractDataCastAdapter abstractDataCastAdapter = (AbstractDataCastAdapter) DefaultAdapters.getAdapter(mappingInfo.getFieldType());
            abstractDataCastAdapter.setWorkBookReaderConfig(workBookReaderConfig);
            adapter = abstractDataCastAdapter;
        }
        CastContext<?> castContext =  new CastContext<>(mappingInfo.getFieldType(),mappingInfo.getFormat());
        if (adapter.support(info.getCellType(),mappingInfo.getFieldType())){
            return adapter.cast(info, castContext);
        }else {
            throw new RuntimeException("不支持转换的类型:["+info.getCellType() +"->"+mappingInfo.getFieldType() +" 字段:["+ mappingInfo.getFieldName() +"]");
        }
    }

    /**
     *
     */
    private void putRowToMapInstance(Map<String,Object> instance, Row row){
        row.cellIterator().forEachRemaining(cell -> {
            int idx = cell.getColumnIndex() + 1;
            String key = "CELL_" + idx;
            instance.put(key,getCellValue(cell, null));
            if (workBookReaderConfig.getReadFeatureAsBoolean(ReadExcelFeature.USE_MAP_DEBUG)){
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
    private CellGetInfo getCellValue(Cell cell, EntityCellMappingInfo mappingInfo){
        // 一般不为null，又map类型传入时，默认使用索引映射
        if (mappingInfo == null){
            mappingInfo = new EntityCellMappingInfo();
            mappingInfo.setColumnPosition(cell.getColumnIndex());
        }
        return switch (mappingInfo.getMappingType()) {
            case INDEX,UNKNOWN -> this.getIndexCellValue(cell, mappingInfo);
            case POSITION -> null;
            // TODO 位置读取
        };
    }

    /**
     * 获取索引映射单元格值
     *
     * @param cell 单元格
     * @param mappingInfo 映射信息
     * @return 单元格值
     */
    private CellGetInfo getIndexCellValue(Cell cell, EntityCellMappingInfo mappingInfo){
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
            switch (cell.getCellType()) {
                case STRING -> value = cell.getStringCellValue();
                case NUMERIC -> value = cell.getNumericCellValue();
                case BOOLEAN -> value = cell.getBooleanCellValue();
                case FORMULA -> value = getFormulaCellValue(cell);
                default -> {
                    LOGGER.error("未知的单元格类型:{},{}",cell.getCellType(), cell);
                }
            };
            return new CellGetInfo(true,value);
        }
    }

    /**
     * 计算公式
     *
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

    /**
     *
     */
    public void reuse(){
        // TODO 用新的配置复用该文件对象
    }
}

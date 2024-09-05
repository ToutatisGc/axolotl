package cn.xvoid.axolotl.excel.reader;

import cn.xvoid.axolotl.Meta;
import cn.xvoid.axolotl.excel.reader.constant.AxolotlDefaultReaderConfig;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.AxolotlAbstractExcelReader;
import cn.xvoid.axolotl.excel.reader.support.AxolotlReadInfo;
import cn.xvoid.axolotl.excel.reader.support.docker.AxolotlCellMapInfo;
import cn.xvoid.axolotl.excel.reader.support.docker.MapDocker;
import cn.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.xvoid.toolkit.clazz.ReflectToolkit;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.toolkit.log.LoggerToolkitKt;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;

import java.io.File;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

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
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,"读取数据错误");
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
                start, this.getRecordRowNumber(),_sheetLevelReaderConfig.getInitialRowPositionOffset()
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
        this.spreadMergedCells(sheet,readerConfig);
        this.readSheetData(sheet,readerConfig,readResult);
        return readResult;
    }

    /**
     * 根据提供的ReaderConfig读取Excel表格数据，并返回AxolotlCellMapInfo对象列表
     * 该方法将结果强制转换为指定类型
     * @see AxolotlCellMapInfo Map映射信息
     * @param readerConfig 用于配置读取操作的配置对象，可以是任意泛型类型
     * @return 返回一个List，其中每个Map代表一行数据，键为字符串，值为AxolotlCellMapInfo对象
     */
    @SuppressWarnings({"unchecked"})
    public List<Map<String, AxolotlCellMapInfo>> readSheetDataAsMapObject(ReaderConfig<?> readerConfig){
        ReaderConfig<?> configedMapReaderConfig = configMapReaderConfig(readerConfig,true);
        return (List<Map<String, AxolotlCellMapInfo>>) this.readSheetData(configedMapReaderConfig);
    }

    /**
     * 将工作表数据读取为平面映射列表
     * Key命名规则为CELL_<单元格索引>[@自定义MapDocker]
     * @see ReaderConfig#setMapDocker(String, MapDocker)
     * @param readerConfig 读取配置，泛型参数表示配置的具体类型
     * @return 返回一个List，每个元素是一个Map，表示一行数据
     */
    @SuppressWarnings({"unchecked"})
    public List<Map<String, Object>> readSheetDataAsFlatMap(ReaderConfig<?> readerConfig){
        ReaderConfig<?> configedMapReaderConfig = configMapReaderConfig(readerConfig,false);
        return (List<Map<String, Object>>) this.readSheetData(configedMapReaderConfig);
    }
    @SuppressWarnings({"unchecked","rawtypes"})
    private ReaderConfig<?> configMapReaderConfig(ReaderConfig<?> readerConfig,boolean convertObjectOrFlat){
        ReaderConfig mapReaderConfig = readerConfig;
        if (readerConfig == null){
            mapReaderConfig = new ReaderConfig<Map<String, AxolotlCellMapInfo>>();
        }
        mapReaderConfig.setCastClass(Map.class);
        mapReaderConfig.setBooleanReadPolicy(ExcelReadPolicy.MAP_CONVERT_INFO_OBJECT,convertObjectOrFlat);
        LoggerToolkitKt.debugWithModule(LOGGER, Meta.MODULE_NAME,"读取策略已被强制设置为[MAP_CONVERT_INFO_OBJECT],并设置读取类型为[Map.class]");
        return mapReaderConfig;
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
            String needRecordInfo = readerConfig.getNeedRecordInfo();
            if (instance != null && needRecordInfo != null ){
                try {
                    Field field = readerConfig.getCastClass().getDeclaredField(needRecordInfo);
                    AxolotlReadInfo axolotlReadInfo = new AxolotlReadInfo();
                    axolotlReadInfo.setSheetIndex(readerConfig.getSheetIndex());
                    axolotlReadInfo.setSheetName(sheet.getSheetName());
                    axolotlReadInfo.setRowNumber(i);
                    ReflectToolkit.setObjectField(instance,field,axolotlReadInfo);
                } catch (NoSuchFieldException e) {
                    throw new RuntimeException(e);
                }
            }
            if (instance!= null){list.add(instance);}
        }
    }

}

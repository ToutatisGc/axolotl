package cn.toutatis.xvoid.axolotl.excel;

import cn.toutatis.xvoid.axolotl.excel.annotations.CellBindProperty;
import cn.toutatis.xvoid.axolotl.excel.annotations.SpecifyCellPosition;
import cn.toutatis.xvoid.axolotl.excel.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.constant.ReadExcelFeature;
import cn.toutatis.xvoid.toolkit.constant.Regex;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;

import static cn.toutatis.xvoid.axolotl.excel.constant.ReadExcelFeature.*;

@ToString
@Getter
@Setter
public class ReaderConfig<T> {

    /**
     * 表索引
     */
    private int sheetIndex = -1;

    /**
     * 表名
     */
    private String sheetName;

    /**
     * 初始行偏移量
     */
    private int initialRowPositionOffset = 0;

    /**
     * 读取的Java类型
     */
    private Class<T> castClass;

    /**
     * 读取的起始行
     * 注意：startIndex = 0时将读取initialRowPositionOffset偏移量之后的行
     */
    private int startIndex = 0;

    /**
     * 读取的结束行
     * endIndex = -1时将读取表尾
     */
    private int endIndex = -1;

    /**
     * 索引映射信息
     * key: 索引
     * value: 映射信息
     */
    private List<EntityCellMappingInfo<?>> indexMappingInfos;

    /**
     * 位置映射信息
     * key: 位置[A5=0,4,B2=1,1]
     */
    private Map<String,EntityCellMappingInfo<?>> positionMappingInfos;

    /**
     * 读取的特性
     */
    private Map<ReadExcelFeature, Object> readFeature = new HashMap<>();

    /**
     *
     */
    public ReaderConfig() {
        this(true);
    }

    /**
     * @param castClass 读取的Java类型
     */
    public ReaderConfig(Class<T> castClass) {
        this(true);
        this.setCastClass(castClass);
    }

    /**
     *
     */
    public ReaderConfig(boolean withDefaultConfig) {
        if (withDefaultConfig) {
            readFeature.putAll(defaultReadFeature());
        }
    }

    /**
     *
     */
    private Map<ReadExcelFeature, Object> defaultReadFeature() {
        Map<ReadExcelFeature, Object> defaultReadFeature = new HashMap<>();
        defaultReadFeature.put(IGNORE_EMPTY_SHEET_ERROR,IGNORE_EMPTY_SHEET_ERROR.getValue());
        defaultReadFeature.put(SORTED_READ_SHEET_DATA,SORTED_READ_SHEET_DATA.getValue());
        defaultReadFeature.put(INCLUDE_EMPTY_ROW,INCLUDE_EMPTY_ROW.getValue());
        defaultReadFeature.put(USE_MAP_DEBUG,USE_MAP_DEBUG.getValue());
        defaultReadFeature.put(DATA_BIND_PRECISE_LOCALIZATION,DATA_BIND_PRECISE_LOCALIZATION.getValue());
        defaultReadFeature.put(CAST_NUMBER_TO_DATE,CAST_NUMBER_TO_DATE.getValue());
        return defaultReadFeature;
    }

    public void setCastClass(Class<T> castClass) {
        this.setCastClass(castClass,false);
    }

    /**
     * @param castClass 读取的Java类型
     */
    public void setCastClass(Class<T> castClass,boolean readClassAnnotation) {
        this.castClass = castClass;
        if (readClassAnnotation) {
            this.processClassAnnotation();
        }
        this.processEntityFieldMappingToCell();
    }

    private void processClassAnnotation() {
        // TODO 读取WorkSheet注解
    }

    /**
     * 处理实体字段映射到单元格
     */
    private void processEntityFieldMappingToCell() {
        Field[] declaredFields = castClass.getDeclaredFields();
        List<EntityCellMappingInfo<?>> entityCellMappingInfos = new ArrayList<>(declaredFields.length);
        HashMap<String, EntityCellMappingInfo<?>> positionMappingInfos = new HashMap<>();
        boolean preciseLocalization = getReadFeatureAsBoolean(DATA_BIND_PRECISE_LOCALIZATION);
        AtomicInteger idx = new AtomicInteger(-1);
        for (Field declaredField : declaredFields) {
            idx.getAndIncrement();
            EntityCellMappingInfo<?> entityCellMappingInfo = new EntityCellMappingInfo<>(declaredField.getType());
            entityCellMappingInfo.setFieldIndex(idx.get());
            entityCellMappingInfo.setFieldName(declaredField.getName());
            // 指定单元格索引
            CellBindProperty cellBindProperty = declaredField.getAnnotation(CellBindProperty.class);
            if (cellBindProperty != null) {
                entityCellMappingInfo.setMappingType(EntityCellMappingInfo.MappingType.INDEX);
                entityCellMappingInfo.setColumnPosition(cellBindProperty.cellIndex());
                entityCellMappingInfo.setDataCastAdapter(cellBindProperty.adapter());
                entityCellMappingInfo.setFormat(cellBindProperty.format());
                entityCellMappingInfos.add(entityCellMappingInfo);
                continue;
            }
            // 指定单元格具体位置
            SpecifyCellPosition specifyCellPosition = declaredField.getAnnotation(SpecifyCellPosition.class);
            if (specifyCellPosition != null) {
                entityCellMappingInfo.setMappingType(EntityCellMappingInfo.MappingType.POSITION);
                entityCellMappingInfo.setFormat(specifyCellPosition.format());
                entityCellMappingInfo.setDataCastAdapter(specifyCellPosition.adapter());
                String position = specifyCellPosition.value().toUpperCase();
                String[] alphaNumeric = Regex.splitAlphaNumeric(position);
                if (alphaNumeric.length == 2) {
                    String columnString = alphaNumeric[0];
                    int columnIndex;
                    boolean bigSheetColumn = columnString.length() > 1;
                    int simpleIdx = columnString.charAt(0) - (int) 'A';
                    if (bigSheetColumn){
                        columnIndex = ((simpleIdx + 1) * 26)+(columnString.charAt(1) - (int) 'A');
                    }else {
                        columnIndex = simpleIdx;
                    }
                    entityCellMappingInfo.setColumnPosition(columnIndex);
                    entityCellMappingInfo.setRowPosition(Integer.parseInt(alphaNumeric[1])-1);
                    positionMappingInfos.put(columnIndex+","+entityCellMappingInfo.getRowPosition(),entityCellMappingInfo);
                }else {
                    throw new IllegalArgumentException("指定单元格位置格式错误");
                }
                continue;
            }
            // 未指定单元格位置默认情况
            entityCellMappingInfo.setMappingType(EntityCellMappingInfo.MappingType.UNKNOWN);
            entityCellMappingInfo.setColumnPosition(preciseLocalization ? -1 : idx.get());
            entityCellMappingInfos.add(entityCellMappingInfo);
        }
        this.positionMappingInfos = positionMappingInfos;
        indexMappingInfos = entityCellMappingInfos;
    }

    /**
     *
     */
    public boolean getReadFeatureAsBoolean(ReadExcelFeature feature) {
        if (feature.getType() != ReadExcelFeature.Type.BOOLEAN){
            throw new IllegalArgumentException("读取特性不是一个布尔类型");
        }
        return readFeature.containsKey(feature) && (boolean) readFeature.get(feature);
    }

    /**
     *
     */
    public void addReadFeature(ReadExcelFeature feature, Object value) {
        readFeature.put(feature, value);
    }

    /**
     * 获取转换类型实例
     */
    @SuppressWarnings("unchecked")
    public T getCastClassInstance(){
        if(castClass!=null){
            try {
                if (castClass == Map.class){
                    if (getReadFeatureAsBoolean(SORTED_READ_SHEET_DATA)){
                        return (T) new LinkedHashMap<String,Object>();
                    }else{
                        return (T) new HashMap<String, Object>();
                    }
                }
                return castClass.getDeclaredConstructor().newInstance();
            } catch (InstantiationException | IllegalAccessException |
                     InvocationTargetException | NoSuchMethodException e) {
                throw new RuntimeException(e);
            }
        }else{
            throw new IllegalArgumentException("转换类型为空");
        }
    }
}

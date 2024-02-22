package cn.toutatis.xvoid.axolotl.excel.reader;

import cn.toutatis.xvoid.axolotl.excel.reader.annotations.*;
import cn.toutatis.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.toutatis.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.toutatis.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.toutatis.xvoid.toolkit.constant.Regex;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import static cn.toutatis.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy.*;

@ToString
@Getter
@Setter
public class ReaderConfig<T> {

    /**
     * 表索引
     * 默认读取第一张表
     */
    private int sheetIndex = 0;

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
     * 读取表为对象
     * 默认读取为List
     */
    private boolean readAsObject = false;

    /**
     * 索引映射信息
     * key: 索引
     * value: 映射信息
     */
    private List<EntityCellMappingInfo<?>> indexMappingInfos;

    /**
     * 位置映射信息
     */
    private List<EntityCellMappingInfo<?>> positionMappingInfos;

    /**
     * 读取的特性
     */
    private Map<ExcelReadPolicy, Object> rowReadPolicyMap = new HashMap<>();

    /**
     * 读取类注解
     */
    private boolean readClassAnnotation = false;

    /**
     * 读取表头最大行数
     * 默认使用默认值10条
     */
    private int searchHeaderMaxRows = -1;

    /**
     * 默认构造
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
            rowReadPolicyMap.putAll(defaultReadPolicy());
        }
    }

    public ReaderConfig(Class<T> castClass, boolean withDefaultConfig) {
        if (withDefaultConfig) {
            rowReadPolicyMap.putAll(defaultReadPolicy());
        }
        this.setCastClass(castClass);

    }

    /**
     *
     */
    private Map<ExcelReadPolicy, Object> defaultReadPolicy() {
        Map<ExcelReadPolicy, Object> defaultReadPolicies = new HashMap<>();
        for (ExcelReadPolicy policy : values()) {
            if (policy.isDefaultPolicy()){
                defaultReadPolicies.put(policy,policy.getValue());
            }
        }
        return defaultReadPolicies;
    }

    /**
     * @param castClass 读取的Java类型
     */
    public void setCastClass(Class<T> castClass) {
        this.setCastClass(castClass,true);
    }

    /**
     * @param castClass 读取的Java类型
     */
    public void setCastClass(Class<T> castClass,boolean readClassAnnotation) {
        // 限制设置类型
        if (List.class.isAssignableFrom(castClass)){
            throw new IllegalArgumentException("请指定一般POJO类型");
        }
        this.castClass = castClass;
        if (readClassAnnotation) {
            this.processClassAnnotation();
        }
        this.processEntityFieldMappingToCell();
    }

    /**
     * 处理实体注解
     * NamingWorkSheet优先级>IndexWorkSheet优先级
     * @see NamingWorkSheet 命名工作表
     * @see IndexWorkSheet 索引指定工作表
     */
    private void processClassAnnotation() {
        NamingWorkSheet namingWorkSheet = castClass.getAnnotation(NamingWorkSheet.class);
        if (namingWorkSheet != null) {
            this.setSheetName(namingWorkSheet.sheetName());
            this.setInitialRowPositionOffset(namingWorkSheet.readRowOffset());
            this.setReadClassAnnotation(true);
            return;
        }
        IndexWorkSheet indexWorkSheet = castClass.getAnnotation(IndexWorkSheet.class);
        if (indexWorkSheet != null) {
            this.setSheetIndex(indexWorkSheet.sheetIndex());
            this.setInitialRowPositionOffset(indexWorkSheet.readRowOffset());
            this.setReadClassAnnotation(true);
        }
    }

    /**
     * 处理实体字段映射到单元格
     * 单元格处理注有具有优先级
     * 指定位置注解优先级>数据绑定注解优先级
     */
    private void processEntityFieldMappingToCell() {
        Field[] declaredFields = castClass.getDeclaredFields();
        List<EntityCellMappingInfo<?>> indexPositionMappingInfos = new ArrayList<>();
        List<EntityCellMappingInfo<?>> positionMappingInfos = new ArrayList<>();
        boolean preciseLocalization = getReadPolicyAsBoolean(DATA_BIND_PRECISE_LOCALIZATION);
        AtomicInteger idx = new AtomicInteger(-1);
        for (Field declaredField : declaredFields) {
            idx.getAndIncrement();
            EntityCellMappingInfo<?> entityCellMappingInfo = new EntityCellMappingInfo<>(declaredField.getType());
            entityCellMappingInfo.setFieldIndex(idx.get());
            entityCellMappingInfo.setFieldName(declaredField.getName());
            // 排除特性
            KeepIntact keepIntact = declaredField.getAnnotation(KeepIntact.class);
            if (keepIntact!= null){
                ExcelReadPolicy[] excludePolicies = keepIntact.excludePolicies();
                entityCellMappingInfo.setExcludePolicies(
                        Arrays.stream(excludePolicies)
                                .collect(Collectors.toMap(policy -> policy, policy -> true))
                );
            }
            // 指定单元格具体位置
            SpecifyPositionBind specifyPositionBind = declaredField.getAnnotation(SpecifyPositionBind.class);
            if (specifyPositionBind != null) {
                entityCellMappingInfo.setMappingType(EntityCellMappingInfo.MappingType.POSITION);
                entityCellMappingInfo.setFormat(specifyPositionBind.format());
                entityCellMappingInfo.setDataCastAdapter(specifyPositionBind.adapter());
                String position = specifyPositionBind.value().toUpperCase();
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
                    positionMappingInfos.add(entityCellMappingInfo);
                }else {
                    throw new IllegalArgumentException("指定单元格位置格式错误");
                }
                continue;
            }
            // 指定单元格索引
            ColumnBind columnBind = declaredField.getAnnotation(ColumnBind.class);
            if (columnBind != null) {
                entityCellMappingInfo.setMappingType(EntityCellMappingInfo.MappingType.INDEX);
                entityCellMappingInfo.setColumnPosition(columnBind.columnIndex());
                String headerName = columnBind.headerName();
                if (Validator.strNotBlank(headerName)){
                    entityCellMappingInfo.setHeaderName(headerName);
                    entityCellMappingInfo.setHeaderNameIndex(columnBind.sameHeaderIdx());
                }
                entityCellMappingInfo.setDataCastAdapter(columnBind.adapter());
                entityCellMappingInfo.setFormat(columnBind.format());
                indexPositionMappingInfos.add(entityCellMappingInfo);
                continue;
            }
            // 未指定单元格位置默认情况
            if (!preciseLocalization){
                entityCellMappingInfo.setMappingType(EntityCellMappingInfo.MappingType.UNKNOWN);
                entityCellMappingInfo.setColumnPosition(idx.get());
                indexPositionMappingInfos.add(entityCellMappingInfo);
            }
        }
        if (positionMappingInfos.isEmpty() && indexPositionMappingInfos.isEmpty()){
            throw new IllegalArgumentException("没有找到任何单元格映射注解");
        }
        this.positionMappingInfos = positionMappingInfos;
        this.indexMappingInfos = indexPositionMappingInfos;
    }

    /**
     * 获取一个布尔值类型的读取策略
     */
    public boolean getReadPolicyAsBoolean(ExcelReadPolicy policy) {
        if (policy.getType() != ExcelReadPolicy.Type.BOOLEAN){
            throw new IllegalArgumentException("读取特性不是一个布尔类型");
        }
        return rowReadPolicyMap.containsKey(policy) && (boolean) rowReadPolicyMap.get(policy);
    }

    /**
     * 设置读取策略
     * @param policy 读取策略
     * @param value 值
     */
    public void setBooleanReadPolicy(ExcelReadPolicy policy, boolean value) {
        if (policy.getType() != ExcelReadPolicy.Type.BOOLEAN){
            throw new IllegalArgumentException("读取特性不是一个布尔类型");
        }
        rowReadPolicyMap.put(policy, value);
    }

    /**
     * 获取转换类型实例
     */
    @SuppressWarnings("unchecked")
    public T getCastClassInstance(){
        if(castClass!=null){
            try {
                if (castClass == Map.class){
                    if (getReadPolicyAsBoolean(SORTED_READ_SHEET_DATA)){
                        return (T) new LinkedHashMap<String,Object>();
                    }else{
                        return (T) new HashMap<String, Object>();
                    }
                }
                return castClass.getDeclaredConstructor().newInstance();
            } catch (InstantiationException | IllegalAccessException |
                     InvocationTargetException | NoSuchMethodException e) {
                throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR,"类型实例化失败:"+e.getMessage());
            }
        }else{
            throw new IllegalArgumentException("转换类型为空");
        }
    }
}

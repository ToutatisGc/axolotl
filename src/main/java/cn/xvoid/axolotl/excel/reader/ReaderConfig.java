package cn.xvoid.axolotl.excel.reader;

import cn.xvoid.axolotl.common.AxolotlCommonConfig;
import cn.xvoid.axolotl.common.annotations.AxolotlDictMapping;
import cn.xvoid.axolotl.excel.reader.annotations.*;
import cn.xvoid.axolotl.excel.reader.constant.EntityCellMappingInfo;
import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.xvoid.axolotl.excel.reader.support.AxolotlReadInfo;
import cn.xvoid.axolotl.excel.reader.support.AxolotlValid;
import cn.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import cn.xvoid.axolotl.toolkit.LoggerHelper;
import cn.xvoid.toolkit.clazz.ReflectToolkit;
import cn.xvoid.toolkit.constant.Regex;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.toolkit.validator.Validator;
import cn.xvoid.axolotl.excel.reader.annotations.*;
import lombok.*;
import org.slf4j.Logger;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

/**
 * 读取配置
 * @param <T> 转换实体
 * @author Toutatis_Gc
 */
@Data
@EqualsAndHashCode(callSuper = true)
public class ReaderConfig<T> extends AxolotlCommonConfig {

    private Logger LOGGER = LoggerToolkit.getLogger(this.getClass());

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
     * 工作表有效列起始范围
     * 默认使用默认值[0,-1]
     * 起始索引为0,结束索引-1为表最后一列
     */
    private int[] sheetColumnEffectiveRange = new int[]{0,-1};

    /**
     * 读取表为对象
     * 默认读取为List
     */
    private boolean readAsObject = false;

    /**
     * 是否需要记录信息
     */
    @Setter(AccessLevel.PROTECTED)
    private String needRecordInfo;
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
     * 验证组
     * @see jakarta.validation.Validator 验证器
     */
    private Class<?>[] validGroups = new Class[]{AxolotlValid.class};

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
     * 设置默认读取策略
     */
    private Map<ExcelReadPolicy, Object> defaultReadPolicy() {
        Map<ExcelReadPolicy, Object> defaultReadPolicies = new HashMap<>();
        for (ExcelReadPolicy policy : ExcelReadPolicy.values()) {
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
            this.setSheetColumnEffectiveRange(namingWorkSheet.sheetColumnEffectiveRange());
            return;
        }
        IndexWorkSheet indexWorkSheet = castClass.getAnnotation(IndexWorkSheet.class);
        if (indexWorkSheet != null) {
            this.setSheetIndex(indexWorkSheet.sheetIndex());
            this.setInitialRowPositionOffset(indexWorkSheet.readRowOffset());
            this.setReadClassAnnotation(true);
            this.setSheetColumnEffectiveRange(indexWorkSheet.sheetColumnEffectiveRange());
        }
        List<Field> allFields = ReflectToolkit.getAllFields(castClass, true);
        Map<ExcelReadPolicy, Object> policyMap = getRowReadPolicyMap();
        if (!policyMap.containsKey(ExcelReadPolicy.SIMPLE_USE_DICT_CODE_TRANSFER)){
            for (Field field : allFields) {
                if (field.getAnnotation(AxolotlDictMapping.class) != null){
                    LoggerHelper.info(LOGGER,"实体发现字典属性，字典映射策略未开启，已自动开启.");
                    this.setBooleanReadPolicy(ExcelReadPolicy.SIMPLE_USE_DICT_CODE_TRANSFER,true);
                    break;
                }
            }
        }
    }

    /**
     * 处理实体字段映射到单元格
     * 单元格处理注有具有优先级
     * 指定位置注解优先级>数据绑定注解优先级
     */
    private void processEntityFieldMappingToCell() {
        List<Field> declaredFields = ReflectToolkit.getAllFields(castClass, true);
        List<EntityCellMappingInfo<?>> indexPositionMappingInfos = new ArrayList<>();
        List<EntityCellMappingInfo<?>> positionMappingInfos = new ArrayList<>();
        boolean preciseLocalization = getReadPolicyAsBoolean(ExcelReadPolicy.DATA_BIND_PRECISE_LOCALIZATION);
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
                entityCellMappingInfo.setDataCastAdapter(columnBind.adapter());
                entityCellMappingInfo.setFormat(columnBind.format());
                String[] headers = columnBind.headerName();
                if (headers.length > 0){
                    indexPositionMappingInfos.addAll(cloneMultipleHeaderMappingInfo(entityCellMappingInfo,headers, columnBind));
                }else {
                    indexPositionMappingInfos.add(entityCellMappingInfo);
                }
                continue;
            }
            if (declaredField.getType() == AxolotlReadInfo.class && this.needRecordInfo == null){
                this.needRecordInfo = declaredField.getName();
                continue;
            }
            // 未指定单元格位置默认情况
            if (!preciseLocalization){
                entityCellMappingInfo.setMappingType(EntityCellMappingInfo.MappingType.UNKNOWN);
                entityCellMappingInfo.setColumnPosition(idx.get());
                indexPositionMappingInfos.add(entityCellMappingInfo);
            }
        }
        if ((castClass != Object.class && castClass != Map.class) &&
                (positionMappingInfos.isEmpty() && indexPositionMappingInfos.isEmpty())){
            throw new IllegalArgumentException(LoggerHelper.format("类[%s]没有找到任何单元格映射注解", castClass.getSimpleName()));
        }
        this.positionMappingInfos = positionMappingInfos;
        this.indexMappingInfos = indexPositionMappingInfos;
    }

    /**
     * 克隆多个头部映射信息。
     * 通过此方法，可以基于一个原始映射信息对象，创建出多个具有相同字段类型但头部名称不同的映射信息对象。
     * 这在处理Excel表格中存在多个相同字段类型但标题名称不同的情况时非常有用。
     *
     * @param original 原始的实体单元格映射信息，用于提供字段类型和基础设置。
     * @param headers 一个字符串数组，包含需要克隆的头部名称。
     * @param columnBind 列绑定信息，用于提供额外的映射配置，如数据转换适配器和格式化信息。
     * @return 返回一个包含克隆的实体单元格映射信息的列表。
     */
    private List<EntityCellMappingInfo<?>> cloneMultipleHeaderMappingInfo(EntityCellMappingInfo<?> original,String[] headers, ColumnBind columnBind){
        List<EntityCellMappingInfo<?>> multipleInfos = new ArrayList<>();
        for (String header : headers) {
            // 创建一个新的实体单元格映射信息对象，使用原始对象的字段类型。
            EntityCellMappingInfo<?> entityCellMappingInfo = new EntityCellMappingInfo<>(original.getFieldType());
            if (Validator.strNotBlank(header)) {
                // 设置克隆对象的头部名称。
                entityCellMappingInfo.setHeaderName(header);
                entityCellMappingInfo.setHeaderNameIndex(columnBind.sameHeaderIdx());
                entityCellMappingInfo.setMappingType(original.getMappingType());
                entityCellMappingInfo.setFieldIndex(original.getFieldIndex());
                entityCellMappingInfo.setFieldName(original.getFieldName());
                entityCellMappingInfo.setColumnPosition(original.getColumnPosition());
                entityCellMappingInfo.setDataCastAdapter(columnBind.adapter());
                entityCellMappingInfo.setFormat(columnBind.format());
                // 将配置完成的克隆对象添加到列表中。
                multipleInfos.add(entityCellMappingInfo);
            } else {
                throw new IllegalArgumentException("类字段 " + original.getFieldName() + " 头部名称不能为空");
            }
        }
        // 返回包含所有克隆对象的列表。
        return multipleInfos;
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
                    if (getReadPolicyAsBoolean(ExcelReadPolicy.SORTED_READ_SHEET_DATA)){
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

    /**
     * 设置列有效范围
     * @param start 开始列位置
     */
    public void setSheetColumnEffectiveRangeStart(int start){
        if (start<0){
            throw new IllegalArgumentException("开始位置不能小于0");
        }
        this.sheetColumnEffectiveRange[0] = start;
    }

    /**
     * 设置列有效范围
     * @param end 结束列位置
     */
    public void setSheetColumnEffectiveRangeEnd(int end){
        this.sheetColumnEffectiveRange[1] = end;
    }

    /**
     * 设置列有效范围
     * @param start 开始列位置
     * @param end 结束列位置
     */
    public void setSheetColumnEffectiveRange(int start,int end){
        this.setSheetColumnEffectiveRangeStart(start);
        this.setSheetColumnEffectiveRangeEnd(end);
    }

    @Override
    public void setDict(int sheetIndex, String field, Map<String, String> dict) {
        if (Validator.strIsBlank(field)){
            throw new IllegalArgumentException("字段不能为空");
        }
        if (dict!= null && !dict.isEmpty()){
            dictionaryMapping.put(sheetIndex,field,dict);
            boolean useDictCode = getReadPolicyAsBoolean(ExcelReadPolicy.SIMPLE_USE_DICT_CODE_TRANSFER);
            if (!useDictCode){
                LoggerHelper.info(LOGGER,"字典映射策略未开启，已自动开启.");
                this.setBooleanReadPolicy(ExcelReadPolicy.SIMPLE_USE_DICT_CODE_TRANSFER,true);
            }
        }
    }
}

package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.components.annotations.SheetTitle;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.support.inverters.DataInverter;
import cn.toutatis.xvoid.axolotl.excel.writer.support.inverters.DefaultDataInverter;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.ExcelWriteThemes;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.SneakyThrows;

import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * 自动写入的写入配置
 * @author Toutatis_Gc
 */
@Data
@EqualsAndHashCode(callSuper = true)
public class AutoWriteConfig extends CommonWriteConfig {

    public AutoWriteConfig() {
        super(true);
        this.calculateColumnIndexes.add(-1);
    }

    public AutoWriteConfig(Class<?> metaClass,boolean withDefaultConfig) {
        super(withDefaultConfig);
        this.metaClass = metaClass;
        setClassMetaData(this.metaClass);
    }

    /**
     * 标题
     */
    private String title;

    /**
     * 元数据类
     */
    private Class<?> metaClass;

    /**
     * 工作表名称
     */
    private String sheetName;

    /**
     * 自定义字体名称
     * 设置此值将使用自定义字体，忽略主题使用字体
     */
    private String fontName;

    /**
     * 样式渲染器
     */
    private ExcelStyleRender styleRender = ExcelWriteThemes.$DEFAULT.getRender();

    /**
     * 数据转换器
     */
    private DataInverter<?> dataInverter = new DefaultDataInverter();

    /**
     * 空值填充字符
     * null值将被填充为空字符串，常用的字符串有"-","未填写","无"
     */
    private String blankValue = "";

    /**
     * 特殊行高映射
     */
    private Map<Integer,Integer> specialRowHeightMapping = new HashMap<>();

    /**
     * 需要计算的列索引
     * -1:为计算所有数字列
     * 0:空为不计算,填充为默认值
     */
    private HashSet<Integer> calculateColumnIndexes = new HashSet<>();

    /**
     * 设置样式渲染器
     * @param styleRender 样式渲染器
     */
    public void setThemeStyleRender(ExcelStyleRender styleRender) {
        this.styleRender = styleRender;
    }

    /**
     * 设置样式渲染器
     * @param theme 主题
     */
    @SneakyThrows
    public void setThemeStyleRender(ExcelWriteThemes theme) {
        this.styleRender = theme.getRender();
    }

    /**
     * 设置样式渲染器
     * @param themeName 主题
     */
    public void setThemeStyleRender(String themeName) {
        this.setThemeStyleRender(ExcelWriteThemes.valueOf(themeName.toUpperCase()));
    }

    /**
     * 设置样式渲染器
     * @param styleRenderClass 主题渲染器
     */
    public void setThemeStyleRender(Class<? extends ExcelStyleRender> styleRenderClass){
        try {
            this.styleRender = styleRenderClass.getDeclaredConstructor().newInstance();
        } catch (InstantiationException | IllegalAccessException | NoSuchMethodException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }
    }


    /**
     * 获取工作表名称
     * @return 工作表名称
     */
    public String getSheetName() {
        if (sheetName == null) {
            return title;
        }
        return sheetName;
    }

    /**
     * 设置元数据类
     * @param metaClass 元数据类
     */
    public void setClassMetaData(Class<?> metaClass){
        if(metaClass != null){
            SheetTitle sheetTitle = metaClass.getDeclaredAnnotation(SheetTitle.class);
            if(sheetTitle != null){
                this.title = sheetTitle.value();
                String sheetName = sheetTitle.sheetName();
                if (Validator.strNotBlank(sheetName)){
                    this.sheetName = sheetName;
                }
            }
        }else {
            throw new AxolotlWriteException("元信息Class为空");
        }
    }

    /**
     * 设置元数据类
     * @param datas 数据列表
     */
    public void setClassMetaData(List<?> datas){
        if(datas != null && !datas.isEmpty()){
            this.setClassMetaData(datas.get(0).getClass());
        }
    }

    /**
     * 添加需要计算的列索引
     * @param index 列索引
     */
    public void addCalculateColumnIndex(int... index){
        Arrays.stream(index).filter(i -> i >= 0).forEach(i -> this.calculateColumnIndexes.add(i));
    }

    /**
     * 添加特殊行高
     * @param row 行
     * @param height 高度
     */
    public void addSpecialRowHeight(int row,int height){
        this.specialRowHeightMapping.put(row,height);
    }
}

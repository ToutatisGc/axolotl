package cn.xvoid.axolotl.excel.writer;

import cn.xvoid.axolotl.excel.writer.components.annotations.SheetTitle;
import cn.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
import cn.xvoid.axolotl.excel.writer.themes.ExcelWriteThemes;
import cn.xvoid.axolotl.exceptions.AxolotlException;
import cn.xvoid.toolkit.log.LoggerToolkit;
import cn.xvoid.toolkit.validator.Validator;
import com.google.common.collect.HashBasedTable;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.SneakyThrows;
import org.slf4j.Logger;

import java.lang.reflect.InvocationTargetException;
import java.util.*;

/**
 * 自动写入的写入配置
 * @author Toutatis_Gc
 */
@Data
@EqualsAndHashCode(callSuper = true)
public class AutoWriteConfig extends CommonWriteConfig {

    private Logger LOGGER = LoggerToolkit.getLogger(this.getClass());

    public AutoWriteConfig() {
        super(true);
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
     * 工作表名称
     */
    private String sheetName;

    /**
     * 自定义字体名称
     * 设置此值将使用自定义字体，忽略主题使用字体
     */
    private String fontName;

    /**
     * 自动列宽比例,
     * 用于配置自动列宽后单元格两端到文字的距离
     * 在启用 ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH (自动列宽)特性 时设置才能生效,
     * 默认为 1.35 保持中文比例显示正常
     */
    private double autoColumnWidthRatio = 1.35D;

    /**
     * 样式渲染器
     */
    private ExcelStyleRender styleRender = ExcelWriteThemes.$DEFAULT.getRender();


    /**
     * 特殊行高映射
     */
    private HashBasedTable<Integer,Integer,Integer> specialRowHeightMapping = HashBasedTable.create();

    /**
     * 需要计算的列索引
     * -1:为计算所有数字列
     * 0:空为不计算,填充为默认值
     */
    private Map<Integer,Set<Integer>> calculateColumnIndexes = new HashMap<>();

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
            AxolotlException axolotlException = new AxolotlException(String.format("创建[%s]渲染器实例失败", styleRenderClass));
            axolotlException.initCause(e);
            throw axolotlException;
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
            this.autoProcessEntity2OpenDictPolicy();
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
     * @param sheetIndex sheet索引
     * @param index 列索引
     */
    public void addCalculateColumnIndex(int sheetIndex,int... index){
        if(index != null){
            Set<Integer> columnIndexes = getCalculateColumnIndexes(sheetIndex);
            for (int i : index) {
                columnIndexes.add(i);
            }
            calculateColumnIndexes.put(sheetIndex,columnIndexes);
        }

    }

    /**
     * 获取需要计算的列索引
     * @param sheetIndex sheet索引
     * @return
     */
    public Set<Integer> getCalculateColumnIndexes(int sheetIndex){
        //默认配置
        Set<Integer> columnIndexes = calculateColumnIndexes.get(sheetIndex);
        if(columnIndexes == null){
            HashSet<Integer> defaultColumnIndexes = new HashSet<>();
            defaultColumnIndexes.add(-1);
            calculateColumnIndexes.put(sheetIndex,defaultColumnIndexes);
        }
        return calculateColumnIndexes.get(sheetIndex);
    }

    /**
     * 添加特殊行高
     * @param sheetIndex sheet索引
     * @param row 行
     * @param height 高度
     */
    public void addSpecialRowHeight(int sheetIndex,int row,int height){
        this.specialRowHeightMapping.put(sheetIndex,row,height);
    }

    /**
     * 添加特殊行高
     * @param sheetIndex sheet索引
     */
    public Map<Integer, Integer> getSpecialRowHeightMapping(int sheetIndex){
        return specialRowHeightMapping.row(sheetIndex);
    }
}

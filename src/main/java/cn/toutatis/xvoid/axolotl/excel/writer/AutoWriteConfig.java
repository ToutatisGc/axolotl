package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.common.annotations.AxolotlDictMapping;
import cn.toutatis.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.components.annotations.SheetTitle;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.CommonWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy;
import cn.toutatis.xvoid.axolotl.excel.writer.support.inverters.DataInverter;
import cn.toutatis.xvoid.axolotl.excel.writer.support.inverters.DefaultDataInverter;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.ExcelWriteThemes;
import cn.toutatis.xvoid.axolotl.toolkit.LoggerHelper;
import cn.toutatis.xvoid.toolkit.clazz.ReflectToolkit;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkit;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import com.google.common.collect.HashBasedTable;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.SneakyThrows;
import org.slf4j.Logger;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;

import static cn.toutatis.xvoid.axolotl.excel.writer.support.base.ExcelWritePolicy.SIMPLE_USE_DICT_CODE_TRANSFER;

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

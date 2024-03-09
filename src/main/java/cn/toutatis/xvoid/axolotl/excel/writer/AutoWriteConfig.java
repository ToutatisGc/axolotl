package cn.toutatis.xvoid.axolotl.excel.writer;

import cn.toutatis.xvoid.axolotl.excel.writer.components.SheetTitle;
import cn.toutatis.xvoid.axolotl.excel.writer.exceptions.AxolotlWriteException;
import cn.toutatis.xvoid.axolotl.excel.writer.style.ExcelStyleRender;
import cn.toutatis.xvoid.axolotl.excel.writer.support.CommonWriteConfig;
import cn.toutatis.xvoid.axolotl.excel.writer.themes.ExcelWriteThemes;
import cn.toutatis.xvoid.toolkit.validator.Validator;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.SneakyThrows;

/**
 * 模板写入配置
 * @author Toutatis_Gc
 */
@Data
@EqualsAndHashCode(callSuper = true)
public class AutoWriteConfig extends CommonWriteConfig {

    public AutoWriteConfig() {
        super(true);
    }

    public AutoWriteConfig(Class<?> metaClass,boolean withDefaultConfig) {
        super(withDefaultConfig);
        this.metaClass = metaClass;
        if(metaClass != null){
            SheetTitle sheetTitle = metaClass.getDeclaredAnnotation(SheetTitle.class);
            if(sheetTitle != null){
                this.title = sheetTitle.value();
                String sheetName = sheetTitle.sheetName();
                if (Validator.strNotBlank(sheetName)){
                    this.sheetName = sheetName;
                }else{
                    this.sheetName = this.title;
                }
            }
        }else {
            throw new AxolotlWriteException("元信息Class为空");
        }

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
     * 样式渲染器
     */
    private ExcelStyleRender styleRender = ExcelWriteThemes.$DEFAULT.getRender();


    public void setStyleRender(ExcelStyleRender styleRender) {
        this.styleRender = styleRender;
    }

    @SneakyThrows
    public void setStyleRender(ExcelWriteThemes theme) {
        this.styleRender = theme.getRender();
    }

    public void setStyleRender(String themeName) {
        this.setStyleRender(ExcelWriteThemes.valueOf(themeName.toUpperCase()));
    }

    public String getSheetName() {
        if (sheetName == null) {
            return title;
        }
        return sheetName;
    }


}

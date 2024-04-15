

![内置色卡](./IndexedColors内置颜色.png)

![填充样式](./FillPatternType填充样式.png)

#### 可配置主题：

​          可配置主题 **AxolotlConfigurableTheme** 支持对标题、表头、内容的单元格进行自定义样式配置，有助于快速开发出满足需求的主题。



##### ConfigurableStyleConfig 接口： 为可配置主题的配置接口，该接口提供了自定义标题、表头、内容等区域的单元格的样式配置方法。

```java

public interface ConfigurableStyleConfig {

    /**
     * 配置全局样式<p>
     * 渲染器初始化时调用 多次写入时，该方法只会被调用一次。<p>
     * 全局样式配置优先级 AutoWriteConfig内样式有关配置 > 此处配置 > 预制值<p>
     * @param cellConfig  样式配置
     */
    default void globalStyleConfig(CellConfigProperty cellConfig){}

    /**
     * 配置程序写入的单元格样式<p>
     * 配置优先级 此处配置 > 全局样式<p>
     * 若要更多精细化的样式配置建议手动插入合计与编号列<p>
     * @param cellConfig  写入策略与对应单元格样式
     */
    @Deprecated
    default void commonStyleConfig(Map<ExcelWritePolicy,CellConfigProperty> cellConfig){}

    /**
     * 配置表头样式（此处为所有表头配置样式，配置单表头样式请在Header对象内配置）<p>
     * 渲染器渲染表头时调用<p>
     * 表头样式配置优先级   Header对象内配置 > 此处配置 > 全局样式<p>
     * @param cellConfig  样式配置
     */
    default void headerStyleConfig(CellConfigProperty cellConfig){}

    /**
     * 配置标题样式（标题是一个整体，此处为整个标题配置样式）<p>
     * 渲染器渲染表头时调用<p>
     * 标题样式配置优先级  此处配置 > 全局样式<p>
     * @param cellConfig  样式配置
     */
    default void titleStyleConfig(CellConfigProperty cellConfig){}

    /**
     * 配置内容样式<p>
     * 渲染内容时，每渲染一个单元格都会调用此方法<p>
     * 内容样式配置优先级  此处配置 > 全局样式<p>
     * @param cellConfig  样式配置
     * @param fieldInfo 单元格内容信息
     */
    default void dataStyleConfig(CellConfigProperty cellConfig, AbstractStyleRender.FieldInfo fieldInfo){}
```



##### 单元格可配置属性 CellConfigProperty：用于传递单元格行高、列宽与单元格常用样式等单元格配置属性

```java
public class CellConfigProperty {

    /**
     * 行高
     */
    private Short rowHeight;

    /**
     * 列宽
     */
    private Short columnWidth;

    /**
     * 单元格水平对齐方式
     */
    private HorizontalAlignment horizontalAlignment;

    /**
     * 单元格对齐方式
     */
    private VerticalAlignment verticalAlignment;

    /**
     * 背景颜色
     */
    private AxolotlColor foregroundColor;

    /**
     * 填充模式
     */
    private FillPatternType fillPatternType;

    /**
     * 边框样式
     */
    private AxolotlCellBorder border;

    /**
     * 字体样式
     */
    private AxolotlCellFont font;


}
```

##### 边框样式配置属性 AxolotlCellBorder：

```java
public class AxolotlCellBorder {
    /**
     * 边框默认样式
     */
    private BorderStyle baseBorderStyle;

    /**
     * 边框默认颜色
     */
    private IndexedColors baseBorderColor;

    /**
     * 上边框样式
     */
    private BorderStyle borderTopStyle;

    /**
     * 上边框颜色
     */
    private IndexedColors topBorderColor;

    /**
     * 下边框样式
     */
    private BorderStyle borderBottomStyle;

    /**
     * 下边框颜色
     */
    private IndexedColors bottomBorderColor;

    /**
     * 左边框样式
     */
    private BorderStyle borderLeftStyle;

    /**
     * 左边框颜色
     */
    private IndexedColors leftBorderColor;

    /**
     * 右边框样式
     */
    private BorderStyle borderRightStyle;

    /**
     * 右边框颜色
     */
    private IndexedColors rightBorderColor;

}
```

##### 字体样式配置属性 AxolotlCellFont：

```java
public class AxolotlCellFont {

    /**
     * 字体名称
     */
    private String fontName;

    /**
     * 是否加粗
     */
    private Boolean bold;

    /**
     * 字体大小
     */
    private Short fontSize;

    /**
     *字体颜色
     */
    private IndexedColors fontColor;

    /**
     *设置文字为斜体
     */
    private Boolean italic;

    /**
     * 使用水平删除线
     */
    private Boolean strikeout;
}
```





##### 预制值：

预制值是**单元格可配置属性**的默认值，若某个**单元格可配置属性**未进行任何自定义配置则会使用预制值填充。

以下是一些**单元格可配置属性**的预制值 ⬇

```
//单元格水平对齐方式
defaultStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
//单元格垂直对齐方式
defaultStyle.setVerticalAlignment(VerticalAlignment.CENTER);
//单元格背景色
defaultStyle.setForegroundColor(StyleHelper.WHITE_COLOR);
//单元格填充方式
defaultStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
//单元格边框样式（四边）
defaultStyle.setBaseBorderStyle(BorderStyle.NONE);
//单元格边框颜色（四边）
defaultStyle.setBaseBorderColor(IndexedColors.BLACK);
//单元格字体名称
defaultStyle.setFontName(StyleHelper.STANDARD_FONT_NAME);
//单元格字体颜色
defaultStyle.setFontColor(IndexedColors.BLACK);
//单元格字体是否加粗
defaultStyle.setBold(false);
//单元格字体大小
defaultStyle.setFontSize(StyleHelper.STANDARD_TEXT_FONT_SIZE);
//单元格字体是否倾斜
defaultStyle.setItalic(false);
//单元格字体是否有删除线
defaultStyle.setStrikeout(false);
```

**部分属性的预制值因生效的单元格所处区域不同而有所区别：**

**行高 rowHeight：**标题：600    表头、内容、程序写入单元格：400  

**列宽 columnWidth：**表头：依据表头单元格的值经过计算得出   内容：继承表头，当没有表头时，指定列宽为 12   程序写入单元格：12

**程序写入的单元格的预制值请参照下文 ->程序写入的单元格样式配置说明**



##### 样式配置生效的优先级与依赖关系：

 * **全局样式配置：**是优先级最低的配置，该配置可用于所有位置的单元格样式配置，若不配置则使用预制值（不同位置的单元格的预制值属性会有所不同）,在**AutoWriteConfig** 内有关样式的配置优先级高于在此处的配置            **优先级**：AutoWriteConfig 内有关样式的配置 > 此处配置 > 预制值

 * **程序写入的单元格样式：**初始配置与全局样式配置一致，在此处配置的属性优先级大于全局样式     **优先级**：此处配置 > 全局样式配置

 * **表头样式配置：**初始配置与全局样式配置一致，在此处配置的属性优先级大于全局样式     在此处配置的表头样式为所有表头单元格的默认样式，因此它的优先级小于直接在表头对象**Header**中进行的样式配置     **优先级**：Header对象内配置 > 此处配置 > 全局样式

 * **内容样式配置：**初始配置与全局样式配置一致，在此处配置的属性优先级大于全局样式     此处的配置**会被读取多次**，每渲染一个内容单元格前都会读取一次配置     **优先级**：此处配置 > 全局样式

   



**关于配置属性有效性的说明：**

 *                    开启 **ExcelWritePolicy.AUTO_CATCH_COLUMN_LENGTH（自动列宽）** 特性后列宽的设置将**不会生效**
 *                    **行高与列宽配置的时效性：**受渲染顺序的影响，后渲染的行高、列宽会**覆盖**掉先渲染的行高、列宽
 *                    **区域的渲染顺序：**表头 -> 标题 -> 内容
 *                    **内容的渲染顺序：**在内容区域的第一行起始依次从左至右渲染单元格，当一个对象的属性或者map的元素渲染完成后换行，进行下一行的渲染
 *                    **程序写入单元格的渲染顺序：**可以将程序写入的单元格当作对应区域内的单元格来确认渲染顺序





**全局样式配置说明：**

​        全局样式配置是最基础的样式配置，它主要为其他的样式配置提供默认配置。在**AutoWriteConfig**中配置的样式属性（如fontName）优先级高于在此处配置的属性，未在此处配置的样式属性会使用**预制值**填充。

**配置方式：**

```java
@Override
public void globalStyleConfig(CellConfigProperty cell) {
    //设置全局的单元格背景色
    cell.setForegroundColor(new AxolotlColor(39,56,86));
    //设置全局行高
    cell.setRowHeight((short) 550);
    //设置全局字体
    AxolotlCellFont axolotlCellFont = new AxolotlCellFont();
    axolotlCellFont.setFontSize((short) 10);
    axolotlCellFont.setFontColor(IndexedColors.WHITE);
    axolotlCellFont.setFontName("微软雅黑");
    cell.setFont(axolotlCellFont);
}
```





##### 标题样式配置说明：

​        此处主要用于控制标题区域内单元格的样式，标题单元格是由多个单元格列合并而来，列宽与表头区域的总宽度相同，因此**不支持配置列宽**。

**配置方式：**

```java
@Override
public void titleStyleConfig(CellConfigProperty cell) {
    //配置标题的行高
    cell.setRowHeight((short) 900);
    //配置标题单元格的背景色
    cell.setForegroundColor(new AxolotlColor( 53,80,125));
    //配置标题单元格的字体
    AxolotlCellFont axolotlCellFont = new AxolotlCellFont();
    axolotlCellFont.setFontSize(StyleHelper.STANDARD_TITLE_FONT_SIZE);
    axolotlCellFont.setBold(true);
    cell.setFont(axolotlCellFont);
}
```





##### 表头样式配置说明：

​        此处主要用于配置表头区域内单元格的通用样式，若需要分别配置每一个表头单元格的样式可以在对应的Header对象中进行指定。

​        配置的行高：为所有表头行的行高（不是总行高，是单个行行高）

​        配置的列宽：为表头底层节点单元格的列宽

**配置方式：**

```java
@Override
public void headerStyleConfig(CellConfigProperty cell) {
    //配置表头底层节点单元格的列宽
    cell.setColumnWidth((short)12);
    //配置表头行行高
    cell.setRowHeight((short) 550);
    //配置表头单元格背景色
    cell.setForegroundColor(new AxolotlColor(34,44,69));
    //配置表头单元格边框样式
    AxolotlCellBorder axolotlCellBorder = new AxolotlCellBorder();
    axolotlCellBorder.setBaseBorderStyle(BorderStyle.NONE);
    axolotlCellBorder.setLeftBorderColor(IndexedColors.BLACK);
    axolotlCellBorder.setBorderLeftStyle(BorderStyle.THIN);
    axolotlCellBorder.setRightBorderColor(IndexedColors.BLACK);
    axolotlCellBorder.setBorderRightStyle(BorderStyle.THIN);
    cell.setBorder(axolotlCellBorder);
}
```





##### 内容样式配置说明：

​        此处主要控制处于内容区域内的单元格样式，每渲染一个单元格前都会重新读取一次配置，同时会传入单元格与渲染内容的相关信息，可根据这些信息对单元格作出不同的样式配置。

##### 单元格内容信息  FieldInfo：

```java
//该类保存了单元格内容的相关信息及单元格的行索引与列索引
public static class FieldInfo{

    /**
     * 属性类型
     */
    private Class<?> clazz;

    /**
     * 属性名称
     */
    private final String fieldName;

    /**
     * 属性值
     */
    private final Object value;

    /**
     * 列索引
     */
    private final int columnIndex;

    /**
     * 行索引
     */
    private final int rowIndex;
}
```

##### 配置方式：

```java
@Override
public void dataStyleConfig(CellConfigProperty cell, AbstractStyleRender.FieldInfo fieldInfo) {
    //设置单元格背景色
    cell.setForegroundColor(new AxolotlColor(39,56,86));
    //设置边框样式
    AxolotlCellBorder axolotlCellBorder = new AxolotlCellBorder();
    axolotlCellBorder.setBaseBorderStyle(BorderStyle.NONE);
    axolotlCellBorder.setTopBorderColor(IndexedColors.BLACK);
    axolotlCellBorder.setBorderTopStyle(BorderStyle.THIN);
    axolotlCellBorder.setBottomBorderColor(IndexedColors.BLACK);
    axolotlCellBorder.setBorderBottomStyle(BorderStyle.THIN);
    //根据该单元格的行索引设置不同的边框样式
    if(fieldInfo.getColumnIndex() == 1){
        cellBorder.setLeftBorderColor(IndexedColors.BLACK);
        cellBorder.setBorderLeftStyle(BorderStyle.MEDIUM);
    }
    //根据该单元格的列索引设置不同的边框样式
    if(fieldInfo.getColumnIndex() == 6){
        axolotlCellBorder.setRightBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderRightStyle(BorderStyle.MEDIUM);
    }
    cell.setBorder(axolotlCellBorder);

}
```





##### （特殊）程序写入的单元格样式配置说明：

​        程序写入的单元格是指由特性配置产生的写入，这些写入是由程序完成的。目前有 ExcelWritePolicy.AUTO_INSERT_SERIAL_NUMBER（自动在第一列插入编号） 和  ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING （自动在结尾插入合计行） 这两种特性可以配置写入的单元格样式。

预制值说明：

 **AUTO_INSERT_SERIAL_NUMBER：**

 *                    表头部分序号单元格行高配置不生效（行高使用**headerStyleConfig**配置的表头样式），样式配置、列宽配置生效，若不配置则所有属性都使用**headerStyleConfig**配置的表头样式
 *                    内容部分的序号单元格可进行样式配置、列宽配置、行高配置(因渲染顺序靠前，行高优先级低，可能会被内容设置的行高覆盖，不建议配置)，若无配置则使用序号单元格后第一个单元格的样式（包括行高、列宽），如不存在任何内容则使用全局样式配置

**AUTO_INSERT_TOTAL_IN_ENDING ：**

 *                    合计行单元格支持行高配置、样式配置
 *                    样式：不配置则自动取上一行单元格的样式   行高：不配置则自动取上一行的行高    列宽：不支持配置，继承上一行

##### 配置方式：

```java
@Override
public void commonStyleConfig(Map<ExcelWritePolicy,CellConfigProperty> cell) {
    //创建一个单元格配置属性类对象
    CellConfigProperty cellConfigProperty = new CellConfigProperty();
    //设置单元格属性...
    AxolotlCellBorder axolotlCellBorder = new AxolotlCellBorder();
    axolotlCellBorder.setBaseBorderStyle(BorderStyle.NONE);
    cellConfigProperty.setBorder(axolotlCellBorder);
    cellConfigProperty.setForegroundColor(new AxolotlColor(255,47,47));
    //将该对象put进形参map Map<ExcelWritePolicy,CellConfigProperty> cell 中，格式是 <特性,配置属性>  
    cell.put(ExcelWritePolicy.AUTO_INSERT_TOTAL_IN_ENDING,cellConfigProperty);
}
```

##### 注意：目前一种特性写入的所有单元格只能支持同一种样式配置。如需更精细化配置建议手动配置写入这些单元格，将他们纳入表头或内容中再进行样式控制。





##### 可配置主题的使用：

1.创建一个类并实现ConfigurableStyleConfig接口，重写对应方法并完成样式配置（配置方法可参考上文的样式配置说明）

```java
public class AxolotlDefaultStyleConfig implements ConfigurableStyleConfig {
    
    //配置内容样式
    @Override
    public void dataStyleConfig(CellConfigProperty cell, AbstractStyleRender.FieldInfo fieldInfo) {
        //设置单元格背景色
        cell.setForegroundColor(new AxolotlColor(39,56,86));
        //设置边框样式
        AxolotlCellBorder axolotlCellBorder = new AxolotlCellBorder();
        axolotlCellBorder.setBaseBorderStyle(BorderStyle.NONE);
        axolotlCellBorder.setTopBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderTopStyle(BorderStyle.THIN);
        axolotlCellBorder.setBottomBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderBottomStyle(BorderStyle.THIN);
        //根据该单元格的行索引设置不同的边框样式
        if(fieldInfo.getColumnIndex() == 1){
            cellBorder.setLeftBorderColor(IndexedColors.BLACK);
            cellBorder.setBorderLeftStyle(BorderStyle.MEDIUM);
        }
        //根据该单元格的列索引设置不同的边框样式
        if(fieldInfo.getColumnIndex() == 6){
            axolotlCellBorder.setRightBorderColor(IndexedColors.BLACK);
            axolotlCellBorder.setBorderRightStyle(BorderStyle.MEDIUM);
        }
        cell.setBorder(axolotlCellBorder);
    }

    //配置表头样式
    @Override
    public void headerStyleConfig(CellConfigProperty cellConfig) {
        //配置表头底层节点单元格的列宽
        cell.setColumnWidth((short)12);
        //配置表头行行高
        cell.setRowHeight((short) 550);
        //配置表头单元格背景色
        cell.setForegroundColor(new AxolotlColor(34,44,69));
        //配置表头单元格边框样式
        AxolotlCellBorder axolotlCellBorder = new AxolotlCellBorder();
        axolotlCellBorder.setBaseBorderStyle(BorderStyle.NONE);
        axolotlCellBorder.setLeftBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderLeftStyle(BorderStyle.THIN);
        axolotlCellBorder.setRightBorderColor(IndexedColors.BLACK);
        axolotlCellBorder.setBorderRightStyle(BorderStyle.THIN);
        cell.setBorder(axolotlCellBorder);
    }
}
```

2.将配置类实例或它的Class类通过有参构造传入AxolotlConfigurableTheme主题中，并把主题配置到AutoWriteConfig中

```java
AutoWriteConfig commonWriteConfig = new AutoWriteConfig();
commonWriteConfig.setThemeStyleRender(new AxolotlConfigurableTheme(new AxolotlDefaultStyleConfig()));
//或者
commonWriteConfig.setThemeStyleRender(new AxolotlConfigurableTheme(AxolotlDefaultStyleConfig.class);
```
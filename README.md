# Axolotl 文档处理框架

![banner](docs.assets/banner.png)

## Part.1 简介

**✨此项目目前处于ALPHA版本✨**

------

​	此项目基于 Apache POI 框架，用于处理文档内容如Excel工作簿等。

​	通过该框架，用户可以轻松读取、写入、以及操作文件中的数据，支持对不同格式的文件进行处理。

​	项目利用 Apache POI 提供的丰富功能，实现了对大型文档的高效处理，并提供了灵活的接口，方便用户根据需求定制化操作。

​	无论是数据导入、导出，还是对内容进行复杂的编辑和分析，此框架都为用户提供了便捷而强大的解决方案，使得文档的处理变得更加高效、灵活。

### 1.1 版本更新说明

#### 🔝 0.0.9-ALPHA-8 更新说明

- 增加对流式读取的支持。**详情查看章节【4.1.7】**
- 对读取器进行抽取方法，增加灵活性。
- 计划支持PDF，添加相关依赖。

#### 🧩历史版本更新说明

请参考文件 [📂ChangeLog变更记录说明](docs.assets/changelog/Index.md)

## Part.2 目前支持功能

|  支持的文件格式   | 目前支持功能 |
| :---------------: | :----------: |
| Excel(.xlsx,.xls) |  导入/导出   |
|     PDF(.pdf)     |  🔜计划支持   |

## Part.3 如何开始使用

### 3.1 添加Maven依赖

**⛔<font color='red'>Java17版本（暂未实装，请使用Java8版本）</font>**

```xml
<dependency>
    <groupId>cn.toutatis</groupId>
    <artifactId>axolotl</artifactId>
    <version>1.0.0-ALPHA</version>
</dependency>
```

**✅<font color='green'>Java8版本</font>**

```xml
<dependency>
    <groupId>cn.toutatis</groupId>
    <artifactId>axolotl</artifactId>
    <version>0.0.8-ALPHA-8</version>
</dependency>
```

### 3.2 Excel 文档操作

#### 3.2.1 读取 Excel文件

📖Excel文件支持类型：

| MIME-TYPE                                                    |         说明          |      | 文件后缀 |
| :----------------------------------------------------------- | :-------------------: | :--: | :------: |
| application/vnd.ms-excel                                     | Excel 97-2003文件版本 |      |  [.xls]  |
| application/vnd.openxmlformats-officedocument.spreadsheetml.sheet | Excel 2007及以上版本  |      | [.xlsx]  |

##### 3.2.1.1 构建文档读取器

```java
// 使用静态方法获取Excel读取器（推荐）
AxolotlExcelReader<Object> excelReader1 = Axolotls.getExcelReader(file);
// 流读取支持
AxolotlExcelReader<Object> excelInsReader = Axolotls.getExcelReader(inputStream);
// 或者使用构造方法创建
AxolotlExcelReader<Object> excelReader2 =new AxolotlExcelReader<>(excelFile);
```

##### 3.1.1.2 读取Excel内容

```java
// 1.打开一个Excel文件
File file = new File("打开一个Excel文件.xlsx");
// 2.新建一个Excel读取器
AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
// 3.读取数据
List<Object> data = excelReader.readSheetData();
System.out.println(data);
```

#### 3.2.2 写入Excel文件

```
// TODO 等待支持
```

## Part.4 详细使用说明

### 4.1 Excel文档读取

🔆在构建AxolotlExcelReader后，**使用readSheetData(args)读取数据**，readSheetData有多种形参方法，详情请查看readSheetData方法源代码，基本上均为readSheetData(ReaderConfig readerConfig)的变种使用。

[📌点击跳转至对应章节](#Anchor-ConfigRead)

🔆框架支持读取Excel为List<T>或者为单个Object实例。

```java
// 将所有实体属性指定为@SpecifyPositionBind可以读取为一个Object
POJO data = reader.readSheetDataAsObject(ReaderConfig readerConfig)
// 读取excel为List数据
List<POJO> data = reader.readSheetData(ReaderConfig readerConfig)
```



#### 4.1.1 注解说明

| 注解（annotations）                   | 用途                                            | 参数说明                                                     |
| ------------------------------------- | ----------------------------------------------- | ------------------------------------------------------------ |
| @IndexWorkSheet                       | [Class]<br />指定具体索引的工作表               | [readRowOffset]读取起始偏移行<br />[sheetIndex]工作表索引[默认值:0] |
| @NamingWorkSheet                      | [Class]<br />指定具体名称的工作表（区分大小写） | [readRowOffset]读取起始偏移行<br />[sheetName]工作表名称     |
| <font color='red'>@ColumnBind*</font> | [Property]<br />实体绑定列位置                  | [columnIndex]列索引<br />[format]日期格式化（数据格式化暂不支持）<br />[adapter]数据适配器<br />[headerName]表头名称<br /> |
| @SpecifyPositionBind                  | [Property]<br />实体绑定具体单元格位置          | [value]单元格位置[举例:A1,B2,C3]<br />[format]日期格式化（数据格式化暂不支持）<br />[adapter]数据适配器 |
| @KeepIntact                           | [Property]<br />赋值字段时排除读取策略          | [excludePolicies]排除的读取策略                              |

🧭POJO示例：

```java
@Data
// 指定索引工作表
@IndexWorkSheet(sheetIndex=0)
public class TestEntity {

    // JSR-303支持
    @NotBlank
    // 读取第0列
    @ColumnBind(columnIndex = 0)
    private String name;

    @Min(value = 1,message = "AAA")
    // 排除特性
    @KeepIntact(excludePolicies = ReadPolicy.CAST_NUMBER_TO_DATE)
    private int age;
    
    // 可以指定列表头对应的列位置
    @ColumnBind(headerName = "备注")
    private String remark;
    
    // 指定具体的单元格位置
    @SpecifyPositionBind("A1")
    private String title;

}
```

#### 4.1.2 迭代器支持

​	读取数据量大时支持迭代器读取。

```java
AxolotlExcelReader<POJO> excelReader = Axolotls.getExcelReader(file, POJO.class);
// 配置表级配置（一般不推荐配置，属于内部维护变量）
excelReader.set_sheetLevelReaderConfig(readerConfig);
while (excelReader.hasNext()){
	List<POJO> next = excelReader.next();
}
```

#### 4.1.3 [重要]以读取配置为参数读取<span id="Anchor-ConfigRead"> </span>

​	一般读取来说，若无特殊读取需求，可以直接构造读取器。

```java
AxolotlExcelReader<TestEntity> reader = Axolotls.getExcelReader(file.getInputStream(), TestEntity.class);
// 根据方法中不同形参读取
List<TestEntity> list = reader.readSheetData(7);
```

​	🔆若是需要更灵活的配置需要构造ReaderConfig配置类。

​	readSheetData的根读取方法为 **<font color="red">readSheetData(readerConfig)</font >**

```java
ReaderConfig<CastClass> readerConfig = new ReaderConfig<>(castClass);
List<TestEntity> list = reader.readSheetData(readerConfig);
```

📖ReaderConfig可配置项：

| 配置项                            | 说明                             | 必填 | 默认值       |
| --------------------------------- | -------------------------------- | ---- | ------------ |
| 构造器(Class<T> castClass)        | 设置读取类                       | 是   | 无           |
| 构造器(boolean withDefaultConfig) | 是否使用默认配置                 | 是   | True         |
| Class<T> castClass                | 设置读取类                       | 是   | 无           |
| sheetIndex                        | 工作表索引                       | 否   | 0            |
| sheetName                         | 工作表名称                       | 否   | 无           |
| initialRowPositionOffset          | 初始行偏移量                     | 否   | 0            |
| startIndex                        | 读取起始行                       | 否   | 0            |
| endIndex                          | 读取结束行                       | 否   | -1（所有）   |
| <del>indexMappingInfos</del>      | 索引映射<br />（一般不用指定）   | 否   | 无           |
| <del>positionMappingInfos</del>   | 单元格映射<br />（一般不用指定） | 否   | 无           |
| rowReadPolicyMap                  | 策略集合                         | 否   | 参考默认策略 |
| searchHeaderMaxRows               | 搜索表头最大行                   | 否   | 10           |

#### 4.1.4 读取策略说明

​	在读取Excel文件数据时，读取到数据时会有默认的读取策略。

​	在使用ReaderConfig时，可以使用以下方法指定读取策略。

```java
// 使用 new ReaderConfig<>(false);可取消所有读取策略，基本上以异常抛出形式作为错误，默认为true，包含以下所有默认读取策略
ReaderConfig<T> readerConfig = new ReaderConfig<>(castClass);
// 使用setBooleanReadPolicy方法可指定读取策略
// 在此实例中，忽略空表异常（IGNORE_EMPTY_SHEET_ERROR）指定为true时，将返回一个空的读取列表，在指定为false时将抛出空表异常
readerConfig.setBooleanReadPolicy(ReadPolicy.IGNORE_EMPTY_SHEET_ERROR, false);
```

📖读取策略说明：

| 读取策略枚举                    | 中文说明                                                     | 策略类型 | 是否为默认策略 | 默认值 |
| ------------------------------- | ------------------------------------------------------------ | -------- | -------------- | ------ |
| IGNORE_EMPTY_SHEET_ERROR        | 忽略空表异常                                                 | Boolean  | true           | true   |
| IGNORE_EMPTY_SHEET_HEADER_ERROR | 忽略空表头的错误                                             | Boolean  | true           | true   |
| INCLUDE_EMPTY_ROW               | 空行也视为有效数据                                           | Boolean  | true           | false  |
| SORTED_READ_SHEET_DATA          | 在使用Map接收时，使用LinkedHashMap                           | Boolean  | true           | true   |
| CAST_NUMBER_TO_DATE             | 判断数字为日期类型将转换为日期格式                           | Boolean  | true           | true   |
| DATA_BIND_PRECISE_LOCALIZATION  | 指定此特性,在按行读取时,若没有指定列名,将不会绑定对象属性<br />否则将按照实体字段顺序自动按照索引绑定数据 | Boolean  | true           | true   |
| TRIM_CELL_VALUE                 | 修整单元格去掉单元格所有的空格和换行符                       | Boolean  | true           | true   |
| USE_MAP_DEBUG                   | 使用Map接收数据时，打印调试信息                              | Boolean  | true           | true   |
| FIELD_EXIST_OVERRIDE            | 如果字段存在值覆盖掉原值                                     | Boolean  | true           | true   |
| **VALIDATE_READ_ROW_DATA**      | 读取数据后校验数据                                           | Boolean  | true           | true   |

#### 4.1.5 JSR-303支持（数据校验）

本框架支持Bean Validation,使用Hibernate-Validator进行实体属性校验。

详情请参考 [📂【Hibernate-Validator 8.0.1.Final 使用手册】](docs.assets\hibernate_validator_reference.pdf) (Java8版本使用为6.2.5.Final)

📖一般POJO类使用注解如下：

| **注解**                    | 说明                                                         |
| --------------------------- | ------------------------------------------------------------ |
| @Nul                        | 被注释的元素必须为 Null                                      |
| @NotNull                    | 被注释的元素必须不为 Null                                    |
| @AssertTrue                 | 被注释的元素必须为 True                                      |
| @AssertFalse                | 被注释的元素必须为 False                                     |
| @Min(value)                 | 被注释的元素必须是一个数字，其值必须大于等于指定的最小值     |
| @Max(value)                 | 被注释的元素必须是一个数字，其值必须小于等于指定的最大值     |
| @DecimalMin(value)          | 被注释的元素必须是一个数字，其值必须大于等于指定的最小值     |
| @DecimalMax(value)          | 被注释的元素必须是一个数字，其值必须小于等于指定的最大值     |
| @Size(max, min)             | 被注释的元素的大小必须在指定的范围内，元素必须为集合，代表集合个数 |
| @Pattern(regexp = )         | 正则表达式校验                                               |
| @Digits (integer, fraction) | 被注释的元素必须是一个数字，其值必须在可接受的范围内         |
| @Past                       | 被注释的元素必须是一个过去的日期                             |
| @Future                     | 被注释的元素必须是一个将来的日期                             |
| @Email                      | 被注释的元素必须是电子邮箱地址                               |
| @Length(min=, max=)         | 被注释的字符串的大小必须在指定的范围内，必须为数组或者字符串，若微数组则表示为数组长度，字符串则表示为字符串长度 |
| @NotEmpty                   | 被注释的字符串的必须非空                                     |
| @Range(min=, max=)          | 被注释的元素必须在合适的范围内                               |
| @NotBlank                   | 被注释的字符串的必须非空                                     |
| @URI                        | 字符串是否是一个有效的URL                                    |

#### 4.1.6 异常处理

​	在读取文件时难免会有读取错误的情况，在发生读取异常时会抛出**AxolotlExcelReadException**来提示异常信息。

**🔆推荐使用为在Web框架使用时向上层抛出异常，在全局异常处理器中统一返回错误信息**

📖AxolotlExcelReadException异常中包含以下内容：

| 可获取内容             | 说明                                   |
| ---------------------- | -------------------------------------- |
| message                | 错误信息                               |
| currentReadRowIndex    | 当前读取行数                           |
| currentReadColumnIndex | 当前读取列数                           |
| humanReadablePosition  | 良好可读性的错误位置（示例：A5单元格） |
| fieldName              | 错误的实体属性                         |
| exceptionType          | 读取错误类型                           |

📖AxolotlExcelReadException.ExceptionType错误类型说明：

| 错误枚举             | 说明                        |
| -------------------- | --------------------------- |
| READ_EXCEL_ERROR     | 读取Excel文件时出现了异常   |
| READ_EXCEL_ROW_ERROR | 读取Excel数据时，出现了异常 |
| CONVERT_FIELD_ERROR  | 转换数据时出现异常          |
| VALIDATION_ERROR     | 校验数据时出现异常          |

#### 4.1.7 StreamReader流读取器支持

​	在读取大的Excel文件（文件大小>=10-16M）时，将文件转换为数据加载进内存时会占用大量的时间和内存，在单个Sheet中数据30w行数据左右时将占用10G内存,时间在1min左右。

​	在读取此类大文件时可以使用 **AxolotlStreamExcelReader** 以流的方式读取数据，减少加载时间和内存占用，该读取器相较于**AxolotlExcelReader** 失去了很多特性，例如获取指定位置数据，分页等。

​	<font color='orange'>**在使用流读取器时只能使用迭代器获取表中数据，并且只能支持xlsx格式。**</font>

```java
// 获取流读取器
AxolotlStreamExcelReader<Object> streamExcelReader = Axolotls.getStreamExcelReader(file);
int recordRowNumber = streamExcelReader.getRecordRowNumber();
System.err.println(recordRowNumber);
// 构建配置（在流读取下表位置，开始位置，结束位置等设置均无用）
ReaderConfig<TestEntity> readerConfig = new ReaderConfig<>(TestEntity.class);
// 获取迭代器
AxolotlExcelStream<TestEntity> dataIterator = streamExcelReader.dataIterator(readerConfig);
int idx = 0;
// 读取数据
while (dataIterator.hasNext()){
	TestEntity entity = dataIterator.next();
	System.out.println(idx+"="+entity);
	idx++;
}
```



### 4.2 Excel文档写入

本框架仅支持XLSX文件写入，性能更优异兼容更好。

> 写入Excel功能将在完整支持后完善

## Part.5 疑难解答

### 📛问题相关

#### 日志框架冲突

日志框架冲突,可将依赖中的日志框架移除。

```basic
SLF4J: Failed to load class "org.slf4j.impl.StaticLoggerBinder".
```

📖官网引用：

> This error is reported when the org.slf4j.impl.StaticLoggerBinder class could not be loaded into memory. This happens when no appropriate SLF4J binding could be found on the class path. Placing one (and only one) of slf4j-nop.jar, slf4j-simple.jar, slf4j-log4j12.jar, slf4j-jdk14.jar or logback-classic.jar on the class path should solve the problem.
>
> 翻译：无法将org.slf4j.impl.StaticLoggerBinder类装入内存。当在类路径上找不到适当的SLF4J绑定时，就会发生这种情况。将slf4j-nop.jar、slf4j-simple.jar、slf4j-log4j12.jar、slf4j-jdk14.jar或logback-classic.jar中的一个(且只有一个)放在类路径上应该可以解决这个问题。

```xml
<dependency>
	<groupId>org.apache.logging.log4j</groupId>
	<artifactId>log4j-api</artifactId>
	<version>2.17.1</version>
</dependency>
<dependency>
	<groupId>cn.toutatis</groupId>
	<artifactId>axolotl</artifactId>
	<version>0.0.8-ALPHA-8</version>
		<exclusions>
            <exclusion>
            	<groupId>org.slf4j</groupId>
                <artifactId>slf4j-log4j12</artifactId>
            </exclusion>
        </exclusions>
</dependency>
```

------

#### 合并单元格内容读取

​	读取<font color='red'>**合并单元格**</font>不同列时会读取到同样的内容，因为本框架**目前**采用的策略为散播策略，会将合并单元格的值散播到合并单元格中各个单元格上（原为合并单元格中第0行，第0列的值），未来如有需要将会把此项作为读取策略作为可配置项。

### 🚸使用疑问

#### @ColumnBind注解中headerName的使用

​	该功能的是为了读取数据时直接按照表头名称读取对应列所设计，解决不同模板之间表头有差异造成读取列错位所设计的功能。
​	指定注解中此参数，会去读取工作表中查找<font color='orange'>**完全匹配**</font>的单元格字符串（例如：备注，地址）所对应的列位置转换为所对应的列索引作为读取列（如果有多个同名表头可指定**sameHeaderIdx**参数区分不同同名列），相当于转化为注解中的columnIndex参数。

<div style="float:right;padding-right:15px">
    提出人：<b>@zhangzk</b> 提出时间：<b>2024-02-03</b>
</div>


------

#### IDEA 引入相关包后import中报错但编译正常

​	出现该问题是由于XVOID包功能由其他语言支持，遇到此问题请升级IDEA到最新版。

<div style="float:right;padding-right:15px">
    提出人：<b>@zongzg</b> 提出时间：<b>2024-02-19</b>
</div>


------



## Part.6 相关链接

[📂Apache POI官方网站](https://poi.apache.org/)

[📂Hibernate Validator官方网站](https://docs.jboss.org/hibernate/stable/validator/reference/en-US/html_single)

[📂Apache PDFBox官方网站](https://pdfbox.apache.org/)


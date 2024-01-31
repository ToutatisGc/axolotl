# Axolotl 文档处理框架

![banner](docs.assets/banner.png)

## 1.简介

​	**✨此项目目前处于ALPHA版本,仅测试使用**✨

​	此项目是一个基于 Apache POI 框架，用于处理 Excel 文档。

​	通过该框架，用户可以轻松读取、写入、以及操作 Excel  文件中的数据，支持对不同格式（xls、xlsx）的文件进行处理。

​	项目利用 Apache POI 提供的丰富功能，实现了对大型 Excel  文档的高效处理，并提供了灵活的接口，方便用户根据需求定制化操作。

​	无论是数据导入、导出，还是对 Excel  内容进行复杂的编辑和分析，此框架都为用户提供了便捷而强大的解决方案，使得 Excel 文档的处理变得更加高效、灵活。

### 1.1 版本更新说明

#### 🔝0.0.5-ALPHA-8 更新说明（New）

- 构造器现在添加对于InputStream的支持
- AxolotlExcelReader.readSheetData读取工作表添加了更多灵活读取方法
- 将读取空行策略默认设置为False（原为True），现在默认不会读取空行内容
- AxolotlExcelWriter将进入Alpha测试

🧩历史版本更新说明

请参考文件 [ChangeLog变更记录说明](docs.assets/changelog/Index.md)

## 2.目前支持功能

|  支持的文件格式   | 目前支持功能 |      |
| :---------------: | :----------: | ---- |
| Excel(.xlsx,.xls) |  导入/导出   |      |
|     PDF(.pdf)     |  🔜计划支持   |      |

## 3.如何开始使用

### 1. 添加Maven依赖

**⛔Java17版本（暂未实装，请使用Java8版本）**

```xml
<dependency>
    <groupId>cn.toutatis</groupId>
    <artifactId>axolotl</artifactId>
    <version>1.0.0-ALPHA</version>
</dependency>
```

**Java8版本**

```xml
<dependency>
    <groupId>cn.toutatis</groupId>
    <artifactId>axolotl</artifactId>
    <version>0.0.5-ALPHA-8</version>
</dependency>
```

### 2. 文档操作

#### 2.1 读取Excel文件

文件支持

##### 2.1.1 构建文档读取器

```java
// 使用静态方法获取Excel读取器（推荐）
AxolotlExcelReader<Object> excelReader1 = Axolotls.getExcelReader(file);
// 流读取支持
AxolotlExcelReader<Object> excelInsReader = Axolotls.getExcelReader(inputStream);
// 或者使用构造方法创建
AxolotlExcelReader<Object> excelReader2 =new AxolotlExcelReader<>(excelFile);
```

### 3.Excel类型读取内容

```java
// 1.打开一个Excel文件
File file = new File("打开一个Excel文件.xlsx");
// 2.新建一个Excel读取器
AxolotlExcelReader<Object> excelReader = Axolotls.getExcelReader(file);
// 3.读取数据
List<Object> data = excelReader.readSheetData();
System.out.println(data);
```

## 4.详细使用说明

#### 1. 注解说明

| 注解（annotations）  | 用途                                            | 参数说明                                                     |
| -------------------- | ----------------------------------------------- | ------------------------------------------------------------ |
| @IndexWorkSheet      | [Class]<br />指定具体索引的工作表               | [sheetIndex]工作表索引[默认值:0]                             |
| @NamingWorkSheet     | [Class]<br />指定具体名称的工作表（区分大小写） | [sheetName]工作表名称                                        |
| @ColumnBind（*）     | [Property]<br />实体绑定列位置                  | [columnIndex]列索引<br />[format]日期格式化（数据格式化暂不支持）<br />[adapter]数据适配器 |
| @SpecifyPositionBind | [Property]<br />实体绑定具体单元格位置          | [value]单元格位置[举例:A1,B2,C3]<br />[format]日期格式化（数据格式化暂不支持）<br />[adapter]数据适配器 |
| @KeepIntact          | [Property]<br />赋值字段时排除读取策略          | [excludePolicies]排除的读取策略                              |

#### 2. 以读取配置为参数读取(推荐方案)



#### 3. 读取策略说明

在读取Excel文件数据时，读取到数据时会有默认的读取策略

在使用ReaderConfig时，可以使用以下方法指定读取策略

```java
// 使用 new ReaderConfig<>(false);可取消所有读取策略，基本上以异常抛出形式作为错误，默认为true，包含以下所有默认读取策略
ReaderConfig<T> readerConfig = new ReaderConfig<>(castClass);
// 使用setBooleanReadFeature方法可指定读取策略
// 在此实例中，忽略空表异常（IGNORE_EMPTY_SHEET_ERROR）指定为true时，将返回一个空的读取列表，在指定为false时将抛出空表异常
readerConfig.setBooleanReadFeature(ReadPolicy.IGNORE_EMPTY_SHEET_ERROR, false);
```

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
| VALIDATE_READ_ROW_DATA          | 读取数据后校验数据                                           | Boolean  | true           | true   |

#### 4. JSR-303支持（数据校验）



## 5. 相关链接

[Apache POI官方网站](https://poi.apache.org/)

[Hibernate Validator官方说明](https://docs.jboss.org/hibernate/stable/validator/reference/en-US/html_single)

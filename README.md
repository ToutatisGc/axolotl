# Axolotl 文档处理工具

![banner](./README.assets/banner.png)

## 1.简介

​	**该项目目前处于ALPHA版本,仅测试使用**

​	这个项目是一个基于 Apache POI 的工具，用于处理 Excel 文档。

​	通过该工具，用户可以轻松读取、写入、以及操作 Excel  文件中的数据，支持对不同格式（xls、xlsx）的文件进行处理。

​	项目利用 Apache POI 提供的丰富功能，实现了对大型 Excel  文档的高效处理，并提供了灵活的接口，方便用户根据需求定制化操作。

​	无论是数据导入、导出，还是对 Excel  内容进行复杂的编辑和分析，这个工具都为用户提供了便捷而强大的解决方案，使得 Excel 文档的处理变得更加高效、灵活。

## 2.目前支持功能

|  支持的文件格式   | 目前支持功能 |
| :---------------: | :----------: |
| Excel(.xlsx,.xls) |     导入     |

## 3.如何开始

### 1.添加Maven依赖

```xml
<dependency>
    <groupId>cn.toutatis</groupId>
    <artifactId>axolotl</artifactId>
    <version>0.0.0-ALPHA</version>
</dependency>
```

### 2.构建文档读取器

```java
// 使用静态方法获取Excel读取器
AxolotlExcelReader<Object> excelReader1 = Axolotls.getExcelReader(file);
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

## 3.详细使用说明

#### 1. 注解说明

| 注解（annotations）  | 用途                               | 参数说明                                                     |
| -------------------- | ---------------------------------- | ------------------------------------------------------------ |
| @IndexWorkSheet      | 指定具体索引的工作表               | [sheetIndex]工作表索引[默认值:0]                             |
| @NamingWorkSheet     | 指定具体名称的工作表（区分大小写） | [sheetName]工作表名称                                        |
| @ColumnBind          | 绑定列位置                         | [columnIndex]列索引<br />[format]日期格式化（数据格式化暂不支持）<br />[adapter]数据适配器 |
| @SpecifyPositionBind | 绑定具体单元格位置                 | [value]单元格位置[举例:A1,B2,C3]<br />[format]日期格式化（数据格式化暂不支持）<br />[adapter]数据适配器 |
| @KeepIntact          | 赋值字段时排除读取策略             | [excludePolicies]排除的读取策略                              |

#### 2.以读取配置为参数读取(推荐方案)



#### 3.读取策略说明

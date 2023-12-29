# Axolotl 文档处理工具

![banner](./README.assets/banner.png)

## 1.简介

**该项目目前处于ALPHA版本,仅测试使用**

该工具可以让你处理文档类型更简单快捷

## 2.支持功能

|  支持的文件格式   | 目前支持功能 |
| :---------------: | :----------: |
| Excel(.xlsx,.xls) |     导入     |

## 3.如何使用

### 1.Excel类型读取内容

```java
// 1.打开一个Excel文件
File file = FileToolkit.getResourceFileAsFile("excel.xlsx");
// 2.新建一个Excel读取器
GracefulExcelReader gracefulExcelReader;
gracefulExcelReader = new GracefulExcelReader(file);
// 3.读取工作Sheet
List<Test> data = gracefulExcelReader.readSheetData(0, Test.class);
// 4.输出读取数据
for (IndexTest map : mapList) {
    System.err.println(map);
}
```


# 更新说明

## 最新更新：

​	最后更新时间：<font color='red'>2024-04-03</font>

​	Java17依赖版本**（LTS）**：<font color='red'>1.0.14</font>

​	Java8依赖版本：<font color='red'>0.0.13-ALPHA</font>

### 🔝Java17 版本 <font color='red'>1.0.14</font> 更新说明

发布时间：[2024-04-03]

- 增加读取特性ALLOW_READ_HIDDEN_SHEET，允许读取隐藏表，否则抛出异常。

- 增加可配置主题。

- 修复写入BUG。

- 增加自动写入器API。

- 完善文档。


## 历史更新：

### 🧩 Java17 版本 <font color='red'>1.0.13</font> 更新说明

发布时间：[2024-03-29]

- 正式发布Java17版本。

- 自动写入功能进入支持阶段。

- 升级VOID-TOOLKIT依赖。

- 统一API命名规则。

- 迁移写入器部分代码结构。

- Java8版本进入补丁修复阶段。

### 🧩 Java8 0.0.10-ALPHA-8 更新说明

- 修复部分API错误。
- 增加指定列范围[sheetColumnEffectiveRange]的ReaderConfig支持。
- 增加默认转换器[support方法]约束。
- 完善使用说明。
- 增加散播策略的读取策略[SPREAD_MERGING_REGION]。
- Excel模板写入进入支持阶段。

### 🧩 Java8 0.0.9-ALPHA-8 更新说明

- 增加对流式读取的支持。**详情查看章节【4.1.7】**
- 对读取器进行抽取方法，增加灵活性。
- 计划支持PDF，添加相关依赖。
- 完善使用说明。

### 🧩 Java8 0.0.8-ALPHA-8 更新说明

- 增加对application/octet-stream的兼容
- 使用@IndexWorkSheet和@NamingWorkSheet的起始读取偏移行支持
- ReaderConfig增加搜索表头最大行的可配置项

### 🧩 Java8 0.0.7-ALPHA-8 更新说明

- 升级Void-Toolkit的依赖
- 更新README.md说明

### 🧩 Java8 0.0.5-ALPHA-8 更新说明

- 构造器现在添加对于InputStream的支持
- AxolotlExcelReader.readSheetData读取工作表添加了更多灵活读取方法
- 将读取空行策略默认设置为False（原为True），现在默认不会读取空行内容
- AxolotlExcelWriter将进入Alpha测试
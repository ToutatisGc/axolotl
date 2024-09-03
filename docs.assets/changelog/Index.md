# 更新说明

## 最新更新：

​	最后更新时间：<font color='red'>2024-09-03</font>

​	Java17依赖版本**（LTS）**：<font color='red'>1.0.17</font>

​	Java8依赖版本：<font color='red'>1.0.17-8</font>

### 🔝 Java17 版本 <font color='red'>1.0.17</font> 更新说明

发布时间：[2024-09]

- 🎉**(重要功能)**新增读取Map拓展功能,可自由添加MapDocker赋予单元格更多类型,详情查看使用文档。
- **默认校验组接口AxolotlValid改为AxolotlValid.Simple内部接口,将接口进一步隔离。**


## 历史更新：

### 🧩 Java17 版本 <font color='red'>1.0.16</font> 更新说明

发布时间：[2024-08-19]

- **(极为重要)修改了整体包的路径引用需要从cn.toutatis.xvoid -> cn.xvoid。**
- **(极为重要)注解校验功能需要在注解的group中添加AxolotlValid组,以避免一个类复用时校验无法通过的问题。**
- 自动序号将从1开始计数。
- 添加AxolotlFaster工具类快捷导出。
- 使用表头匹配功能时,增加了移除单元格左右空格的优化，并允许多个表头名称去匹配,允许相对模糊的匹配。
- 修复了一些特性引起的读取逻辑不清晰。

### 🧩 Java17 版本 <font color='red'>1.0.15</font> 更新说明

发布时间：[2024-06-27]

- 增加读取特性READ_FIELD_USE_SETTER，读取字段使用字段的Setter方法，而并非使用反射赋值。
- 增加字典映射支持，使用config.setDict()方法，可设置字典码与字面量的转换。
- 增加写入特性SIMPLE_USE_GETTER_METHOD写入时，获取值使用Getter方法。
- 为模板写入器增加了数据转换器的支持，同自动写入器保持一致。
- 新增`AxolotlFaster`类添加语法糖，使导入导出时减少配置方法语法更精简。
- 修复部分多表写入BUG。

### 🧩 Java17 版本 1.0.14 更新说明

发布时间：[2024-04-03]

- 增加读取特性ALLOW_READ_HIDDEN_SHEET，允许读取隐藏表，否则抛出异常。

- 增加可配置主题。

- 修复写入BUG。

- 增加自动写入器API。

- 完善文档。

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
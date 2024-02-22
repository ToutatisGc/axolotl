### 🔝0.0.9-ALPHA-8 更新说明

- 增加对流式读取的支持。**详情查看章节【4.1.7】**
- 对读取器进行抽取方法，增加灵活性。
- 计划支持PDF，添加相关依赖。
- 完善使用说明。

### 🧩 0.0.8-ALPHA-8 更新说明

- 增加对application/octet-stream的兼容
- 使用@IndexWorkSheet和@NamingWorkSheet的起始读取偏移行支持
- ReaderConfig增加搜索表头最大行的可配置项

### 🧩 0.0.7-ALPHA-8 更新说明

- 升级Void-Toolkit的依赖
- 更新README.md说明

### 🧩 0.0.5-ALPHA-8 更新说明

- 构造器现在添加对于InputStream的支持
- AxolotlExcelReader.readSheetData读取工作表添加了更多灵活读取方法
- 将读取空行策略默认设置为False（原为True），现在默认不会读取空行内容
- AxolotlExcelWriter将进入Alpha测试
package cn.xvoid.axolotl.excel.reader;

import cn.xvoid.axolotl.excel.reader.constant.ExcelReadPolicy;

/**
 * 读取配置构建器
 *
 * @param <T> 泛型类型
 */
public class ReadConfigBuilder<T> {

    private final ReaderConfig<T> readerConfig;

    /**
     * 构造函数，接受一个Class类型参数用于泛型类型的指定
     *
     * @param castClass 类型
     */
    public ReadConfigBuilder(Class<T> castClass) {
        readerConfig = new ReaderConfig<>(castClass,true);
    }

    /**
     * 构造方法
     *
     * @param castClass         类型
     * @param withDefaultConfig 是否使用默认配置
     */
    public ReadConfigBuilder(Class<T> castClass, boolean withDefaultConfig) {
        readerConfig = new ReaderConfig<>(castClass, withDefaultConfig);
    }

    /**
     * 设置工作表索引
     *
     * @param sheetIndex 工作表索引
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setSheetIndex(int sheetIndex) {
        readerConfig.setSheetIndex(sheetIndex);
        return this;
    }

    /**
     * 设置工作表名称
     *
     * @param sheetName 工作表名称
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setSheetName(String sheetName) {
        readerConfig.setSheetName(sheetName);
        return this;
    }

    /**
     * 设置起始行索引
     *
     * @param startIndex 起始行索引
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setStartIndex(int startIndex) {
        readerConfig.setStartIndex(startIndex);
        return this;
    }

    /**
     * 设置结束行索引
     *
     * @param endIndex 结束行索引
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setEndIndex(int endIndex) {
        readerConfig.setEndIndex(endIndex);
        return this;
    }

    /**
     * 设置初始行位置的偏移量
     *
     * @param initialRowPositionOffset 偏移量值
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setInitialRowPositionOffset(int initialRowPositionOffset) {
        readerConfig.setInitialRowPositionOffset(initialRowPositionOffset);
        return this;
    }

    /**
     * 设置起始行索引和结束行索引
     *
     * @param start 起始行索引
     * @param end   结束行索引
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setStartIndexAndEndIndex(int start, int end) {
        readerConfig.setStartIndex(start);
        readerConfig.setEndIndex(end);
        return this;
    }

    /**
     * 设置读取头的最大行数
     * @param maxRows 最大行数
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setSearchHeaderMaxRows(int maxRows) {
        readerConfig.setSearchHeaderMaxRows(maxRows);
        return this;
    }

    /**
     * 设置读取范围起始列索引
     * @param start 起始列索引
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setSheetColumnEffectiveRangeStart(int start) {
        readerConfig.setSheetColumnEffectiveRangeStart(start);
        return this;
    }

    /**
     * 设置读取范围结束列索引
     * @param end 结束列索引
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setSheetColumnEffectiveRangeEnd(int end) {
        readerConfig.setSheetColumnEffectiveRangeEnd(end);
        return this;
    }

    /**
     * 设置读取范围起始列索引和结束列索引
     * @param start 起始列索引
     * @param end 结束列索引
     * @return 返回ReadConfigBuilder对象
     */
    public ReadConfigBuilder<T> setSheetColumnEffectiveRange(int start,int end) {
        readerConfig.setSheetColumnEffectiveRange(start,end);
        return this;
    }

    /**
     * 设置Excel读取策略的配置方法
     *
     * @param policy Excel读取策略，用于指定特定的读取行为或选项
     * @param value 与读取策略关联的布尔值，用于启用或禁用特定行为
     * @return 返回配置构建器对象，以支持链式调用
     * 本方法通过接收一个Excel读取策略（ExcelReadPolicy）和一个布尔值，来配置Excel文件的读取方式
     * ExcelReadPolicy是一个枚举类型，包含了不同读取策略的选项，如是否自动过滤空行、是否忽略错误的公式等
     * 通过此方法，可以在构建Excel读取配置的过程中，灵活地设置各种读取策略，以满足不同的需求
     */
    public ReadConfigBuilder<T> setBooleanReadPolicy(ExcelReadPolicy policy, boolean value) {
        readerConfig.setBooleanReadPolicy(policy,value);
        return this;
    }


    /**
     * 构建读取配置
     *
     * @return 返回ReaderConfig对象
     */
    public ReaderConfig<T> build() {
        return readerConfig;
    }

}

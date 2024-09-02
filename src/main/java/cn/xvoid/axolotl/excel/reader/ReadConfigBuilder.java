package cn.xvoid.axolotl.excel.reader;

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
     * 构建读取配置
     *
     * @return 返回ReaderConfig对象
     */
    public ReaderConfig<T> build() {
        return readerConfig;
    }

}

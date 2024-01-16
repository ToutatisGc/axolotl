package cn.toutatis.xvoid.axolotl.excel;

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
        readerConfig = new ReaderConfig<>(castClass);
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
     * 构建读取配置
     *
     * @return 返回ReaderConfig对象
     */
    public ReaderConfig<T> build() {
        return readerConfig;
    }

}

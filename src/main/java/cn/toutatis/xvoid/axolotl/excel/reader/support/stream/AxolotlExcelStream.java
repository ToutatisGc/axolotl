package cn.toutatis.xvoid.axolotl.excel.reader.support.stream;

import cn.toutatis.xvoid.axolotl.excel.reader.AxolotlStreamExcelReader;
import cn.toutatis.xvoid.axolotl.excel.reader.ReaderConfig;
import org.apache.poi.ss.usermodel.Row;

import java.util.Iterator;

/**
 * 流式Sheet读取迭代器
 * @param <T> 转换实体泛型
 */
public class AxolotlExcelStream<T> implements Iterator<T> {

    /**
     * 行迭代器
     */
    private final Iterator<Row> rowIterator;

    /**
     * 读取配置
     */
    private final ReaderConfig<T> readerConfig;

    /**
     * 流读取器
     */
    private final AxolotlStreamExcelReader<?> reader;

    public AxolotlExcelStream(AxolotlStreamExcelReader<?> reader, ReaderConfig<T> readerConfig) {
        this.readerConfig = readerConfig;
        this.reader = reader;
        this.rowIterator = reader.getWorkBookContext().getIndexSheet(readerConfig.getSheetIndex()).rowIterator();
    }

    /**
     * 是否有下一行
     * @return 是否有下一行
     */
    @Override
    public boolean hasNext() {
        return rowIterator.hasNext();
    }

    /**
     * 获取下一行并转换为实体
     * @return 转换实体
     */
    @Override
    public T next() {
        return reader.castRow2Instance(rowIterator.next(), readerConfig);
    }
}

package cn.xvoid.axolotl.excel.reader.support.stream;

import cn.xvoid.axolotl.excel.reader.AxolotlStreamExcelReader;
import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.support.exceptions.AxolotlExcelReadException;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

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
     * 迭代器方式获取数据<p>
     * 是否有下一行<p>
     * 该方法用于获取全部数据，起始行、结束行、初始行偏移量等读取范围配置皆不生效 <p>
     * 若要使 起始行、结束行、初始行偏移量 等范围配置生效，请使用 {@link AxolotlExcelStream#readDataBatch(int, ReadBatchTask<T>) }来获取数据
     * @return 是否有下一行
     */
    @Override
    public boolean hasNext() {
        return rowIterator.hasNext();
    }

    /**
     * 迭代器方式获取数据<p>
     * 获取下一行并转换为实体<p>
     * 该方法用于获取全部数据，起始行、结束行、初始行偏移量等读取范围配置皆不生效 <p>
     * 若要使 起始行、结束行、初始行偏移量 等范围配置生效，请使用 {@link AxolotlExcelStream#readDataBatch(int, ReadBatchTask<T>) }来获取数据
     * @return 转换实体
     */
    @Override
    public T next() {
        return reader.castRow2Instance(rowIterator.next(), readerConfig);
    }

    /**
     * 按行获取数据，使用范围配置进行数据过滤
     * @return 一行数据
     */
    private Object nextObject() {
        Row next = rowIterator.next();
        int startIndex = readerConfig.getStartIndex();
        int endIndex = readerConfig.getEndIndex();
        if (startIndex == 0){
            int initialRowPositionOffset = readerConfig.getInitialRowPositionOffset();
            if (initialRowPositionOffset > 0){
                startIndex = startIndex + initialRowPositionOffset;
            }
        }

        if (next.getRowNum() < startIndex){
            return new SpecifyRow();
        }
        if (next.getRowNum() >= endIndex && endIndex != -1){
            return new SpecifyRow();
        }
        return reader.castRow2Instance(next, readerConfig);
    }


    /**
     * 因超出读取范围被遗弃的行
     */
    public static class SpecifyRow{

    }


    /**
     * 分批次读取流中的数据 <p>
     * 用此方法获取数据时，起始行、结束行、初始行偏移量 等范围配置生效
     * @param batchSize 每批数据的数量
     * @param task 读取任务  每批数据读取结束时都会执行任务一次
     */
    public void readDataBatch(int batchSize, ReadBatchTask<T> task){
        if(batchSize < 1){
            throw new IllegalArgumentException("每批数据的数量必须为正整数");
        }
        List<T> data = new ArrayList<>();
        int idx = 0;
        while (this.hasNext()){
            Object next = this.nextObject();
            if (next instanceof SpecifyRow){
                continue;
            }
            idx++;
            @SuppressWarnings("unchecked")
            T entity = (T) next;
            data.add(entity);
            if(idx == batchSize){
                task.execute(new ArrayList<>(data));
                data.clear();
                idx = 0;
            }
        }
        if(!data.isEmpty()){
            task.execute(new ArrayList<>(data));
        }
    }
}

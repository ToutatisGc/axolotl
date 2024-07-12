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
public class AxolotlExcelStream<T> implements Iterator<Object> {

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
     * 是否有下一行<p>
     * 该方法无法直接获取到正确的读取数据，请使用 {@link AxolotlExcelStream#readDataBatch(int, ReadBatchTask<T>) }来获取流数据
     * @return 是否有下一行
     */
    @Override
    @Deprecated
    public boolean hasNext() {
        return rowIterator.hasNext();
    }

    /**
     * 获取下一行并转换为实体<p>
     * 该方法无法直接获取到正确的读取数据，请使用 {@link AxolotlExcelStream#readDataBatch(int, ReadBatchTask<T>) }来获取流数据
     * @return 转换实体
     */
    @Override
    @Deprecated
    public Object next() {
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

    /*@SuppressWarnings("unchecked")
    public T nextObject(){
        Object next = next();
        if (next instanceof SpecifyRow){
            return null;
        }
        return (T) next;
    }
*/

    /**
     * 因超出读取范围被遗弃的行
     */
    public static class SpecifyRow{

    }


    /**
     * 分批次读取流中的数据
     * @param batchSize 每批数据的数量
     * @param task 读取任务  每批数据读取结束时都会执行任务一次
     */
    public void readDataBatch(int batchSize, ReadBatchTask<T> task){
        if(batchSize < 1){
            throw new AxolotlExcelReadException(AxolotlExcelReadException.ExceptionType.READ_EXCEL_ERROR, "每批数据的数量必须为正整数");
        }
        List<T> data = new ArrayList<>();
        int idx = 0;
        while (this.hasNext()){
            Object next = this.next();
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

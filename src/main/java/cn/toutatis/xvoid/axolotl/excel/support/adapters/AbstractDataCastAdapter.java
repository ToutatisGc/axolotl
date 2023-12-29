package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.ReaderConfig;
import lombok.Getter;

@Getter
public abstract class AbstractDataCastAdapter<T,FT> implements DataCastAdapter<FT> {

    private ReaderConfig<T> readerConfig;

    public void setReaderConfig(ReaderConfig<T> readerConfig) {
        this.readerConfig = readerConfig;
    }

}

package cn.toutatis.xvoid.axolotl.excel.support.adapters;

import cn.toutatis.xvoid.axolotl.excel.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.excel.support.WorkBookReaderConfig;
import lombok.Getter;

@Getter
public abstract class AbstractDataCastAdapter<T,FT> implements DataCastAdapter<FT> {

    private WorkBookReaderConfig<T> workBookReaderConfig;

    public void setWorkBookReaderConfig(WorkBookReaderConfig<T> workBookReaderConfig) {
        this.workBookReaderConfig = workBookReaderConfig;
    }

}

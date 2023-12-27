package cn.toutatis.xvoid.axolotl.support.adapters;

import cn.toutatis.xvoid.axolotl.support.DataCastAdapter;
import cn.toutatis.xvoid.axolotl.support.WorkBookReaderConfig;
import lombok.Getter;

@Getter
public abstract class AbstractDataCastAdapter<T> implements DataCastAdapter<T> {

    private WorkBookReaderConfig<T> workBookReaderConfig;

    public void setWorkBookReaderConfig(WorkBookReaderConfig<T> workBookReaderConfig) {
        this.workBookReaderConfig = workBookReaderConfig;
    }

}

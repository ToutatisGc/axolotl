package cn.xvoid.axolotl.excel.reader.support.docker;

import lombok.Data;

@Data
public abstract class AbstractMapDocker<T> implements MapDocker<T> {

//    private String suffix;

    private Boolean nullDisplay;



}

package cn.xvoid.axolotl.excel.reader.support.docker.impl;

import cn.xvoid.axolotl.excel.reader.ReaderConfig;
import cn.xvoid.axolotl.excel.reader.support.CellGetInfo;
import cn.xvoid.axolotl.excel.reader.support.docker.AbstractMapDocker;

public class PlainTextMapDocker extends AbstractMapDocker<String> {

    public static final String SUFFIX_NAME = "PLAIN_TEXT";

    @Override
    public String convert(int index, CellGetInfo cellGetInfo, ReaderConfig<?> readerConfig) {
        return null;
    }
}

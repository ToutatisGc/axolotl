package cn.toutatis.xvoid.axolotl.common;

import lombok.Data;

/**
 * 文档操作结果
 * 主要用于产生对文档操作的结果记录
 * 用户在操作文件时, 可以通过该结果记录来判断文件操作是否成功
 * @author Toutatis_Gc
 */
@Data
public class AxolotlResult {

    private String progressId;

}

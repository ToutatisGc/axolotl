package cn.toutatis.xvoid.axolotl.toolkit;

import cn.toutatis.xvoid.axolotl.Meta;
import cn.toutatis.xvoid.toolkit.log.LoggerToolkitKt;
import org.slf4j.Logger;

/**
 * Logger辅助类
 */
public class LoggerHelper {

    /**
     * 格式化字符串
     * @param message 格式化字符串
     * @param args 格式化参数
     * @return 格式化后的字符串
     */
    public static String format(String message,Object... args){
        /*JDK17*/
        /*return message.formatted(args);*/
        /*JDK8*/
        return String.format(message,args);
    }

    public static void debug(Logger logger,String message) {
        LoggerToolkitKt.debugWithModule(logger, Meta.MODULE_NAME,message);
    }

    public static void debug(Logger logger,String message,Object... args) {
        LoggerToolkitKt.debugWithModule(logger, Meta.MODULE_NAME,format(message, args));
    }

    public static void info(Logger logger,String message){
        LoggerToolkitKt.infoWithModule(logger,Meta.MODULE_NAME,message);
    }

    public static void info(Logger logger,String message,Object... args){
        LoggerToolkitKt.infoWithModule(logger,Meta.MODULE_NAME,format(message, args));
    }

    public static void warn(Logger logger,String message){
        LoggerToolkitKt.warnWithModule(logger,Meta.MODULE_NAME,message);
    }

    public static void warn(Logger logger,String message,Object... args){
        LoggerToolkitKt.warnWithModule(logger,Meta.MODULE_NAME,format(message, args));
    }

    public static void error(Logger logger,String message){
        LoggerToolkitKt.errorWithModule(logger,Meta.MODULE_NAME,message);
    }

    public static void error(Logger logger,String message,Object... args){
        LoggerToolkitKt.errorWithModule(logger,Meta.MODULE_NAME,format(message, args));
    }
}

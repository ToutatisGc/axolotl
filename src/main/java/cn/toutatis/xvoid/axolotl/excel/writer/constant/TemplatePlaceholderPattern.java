package cn.toutatis.xvoid.axolotl.excel.writer.constant;

import java.util.regex.Pattern;

/**
 * Excel模板占位符正则表达式
 * @author Toutatis_Gc
 */
public class TemplatePlaceholderPattern {

    /**
     * 单次引用占位符正则表达式
     */
    public static final String SINGLE_REFERENCE_TEMPLATE_PLACEHOLDER = "\\$\\{([^}]*)\\}";
    public static final Pattern SINGLE_REFERENCE_TEMPLATE_PATTERN = Pattern.compile(SINGLE_REFERENCE_TEMPLATE_PLACEHOLDER);

    /**
     * 循环引用占位符正则表达式
     */
    public static final String CIRCLE_REFERENCE_TEMPLATE_PLACEHOLDER = "#\\{([^}]*)\\}";
    public static final Pattern CIRCLE_REFERENCE_TEMPLATE_PATTERN = Pattern.compile(CIRCLE_REFERENCE_TEMPLATE_PLACEHOLDER);

}

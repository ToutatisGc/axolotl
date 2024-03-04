package cn.toutatis.xvoid.axolotl.excel.writer.constant;

import java.util.regex.Pattern;

/**
 * Excel模板占位符正则表达式
 * @author Toutatis_Gc
 */
public class TemplatePlaceholderPattern {

    public static final String STANDARD_TEMPLATE_PLACEHOLDER = "\\{([^}]*)\\}";

    /**
     * 单次引用占位符正则表达式
     */
    public static final String SINGLE_REFERENCE_TEMPLATE_PLACEHOLDER = "\\$"+STANDARD_TEMPLATE_PLACEHOLDER;
    public static final Pattern SINGLE_REFERENCE_TEMPLATE_PATTERN = Pattern.compile(SINGLE_REFERENCE_TEMPLATE_PLACEHOLDER);

    /**
     * 循环引用占位符正则表达式
     */
    public static final String CIRCLE_REFERENCE_TEMPLATE_PLACEHOLDER = "#"+STANDARD_TEMPLATE_PLACEHOLDER;
    public static final Pattern CIRCLE_REFERENCE_TEMPLATE_PATTERN = Pattern.compile(CIRCLE_REFERENCE_TEMPLATE_PLACEHOLDER);

    /**
     * 合计占位符正则表达式
     */
    public static final String AGGREGATE_REFERENCE_TEMPLATE_PLACEHOLDER = "&"+STANDARD_TEMPLATE_PLACEHOLDER;
    public static final Pattern AGGREGATE_REFERENCE_TEMPLATE_PATTERN = Pattern.compile(AGGREGATE_REFERENCE_TEMPLATE_PLACEHOLDER);

}

package cn.wangxs.excel;

import org.apache.poi.hssf.util.HSSFColor;

import java.lang.reflect.Method;

/**
 * 字段信息
 *
 * @author wangxuesong
 */
public class FieldInfo {

    /**
     * 表头
     */
    private String name;
    /**
     * 顺序
     */
    private int order;
    /**
     * 格式
     */
    private String format;
    /**
     * 宽度
     */
    private int width;
    /**
     * 默认值
     */
    private String defaultValue;
    /**
     * 标记
     */
    private int[] tags;
    /**
     * 合并到
     */
    private String mergeTo;
    /**
     * 分隔符
     */
    private String separator;
    /**
     * 处理"第3周"类似情况
     */
    private String string;
    /**
     * 满足表达式后的颜色
     */
    private HSSFColor color;
    /**
     * 需要高亮显示的表达式
     */
    private String expression;

    /**
     * 字段名称
     */
    private String filedName;

    /**
     * 字段的getter 方法
     */
    private Method method;

    public FieldInfo() {

    }

    public FieldInfo(String name, int order, String format, int width, String defaultValue, Method method, String
            mergeTo, String separator, String string, int[] tags, Class<HSSFColor> color, String expression, String filedName) {
        this.name = name;
        this.order = order;
        this.format = format;
        this.width = width;
        this.method = method;
        this.defaultValue = defaultValue;
        this.mergeTo = mergeTo;
        this.separator = separator;
        this.string = string;
        this.tags = tags;
        this.color = getInstance(color);
        this.expression = expression;
        this.filedName = filedName;
    }

    private HSSFColor getInstance(Class colorClassName) {
        Object o = null;
        try {
            o = colorClassName.newInstance();
        } catch (Exception e) {
            e.printStackTrace();
        }
        if (o instanceof HSSFColor) {
            return (HSSFColor) o;
        }

        return new HSSFColor.RED();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getOrder() {
        return order;
    }

    public void setOrder(int order) {
        this.order = order;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public int[] getTags() {
        return tags;
    }

    public void setTags(int[] tags) {
        this.tags = tags;
    }

    public Method getMethod() {
        return method;
    }

    public void setMethod(Method method) {
        this.method = method;
    }

    public String getDefaultValue() {
        return defaultValue;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    public String getMergeTo() {
        return mergeTo;
    }

    public void setMergeTo(String mergeTo) {
        this.mergeTo = mergeTo;
    }

    public String getSeparator() {
        return separator;
    }

    public void setSeparator(String separator) {
        this.separator = separator;
    }

    public String getString() {
        return string;
    }

    public void setString(String string) {
        this.string = string;
    }

    public String getExpression() {
        return expression;
    }

    public void setExpression(String expression) {
        this.expression = expression;
    }

    public HSSFColor getColor() {
        return color;
    }

    public void setColor(HSSFColor color) {
        this.color = color;
    }

    public String getFiledName() {
        return filedName;
    }

    public void setFiledName(String filedName) {
        this.filedName = filedName;
    }
}

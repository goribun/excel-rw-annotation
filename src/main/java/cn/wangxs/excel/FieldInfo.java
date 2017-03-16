package cn.wangxs.excel;

import java.lang.reflect.Method;

/**
 * 字段信息
 *
 * @author wangxuesong
 */
public class FieldInfo {

    private String name;//表头
    private int order;//顺序
    private String format;//格式
    private int width;//宽度
    private String defaultValue;//默认值
    private int[] tags;//标记
    private String mergeTo;//合并到
    private String separator;//分隔符
    private String string;//处理"第3周"类似情况

    private Method method;//getter

    public FieldInfo() {

    }

    public FieldInfo(String name, int order, String format, int width, String defaultValue, Method method, String
            mergeTo, String separator, String string, int[] tags) {
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

}

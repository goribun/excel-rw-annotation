package cn.wangxs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel导出/导入字段注解
 *
 * @author wangxuesong
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelField {

    /**
     * 列名
     */
    String name() default "";

    /**
     * 格式（作用于日期和小数点）
     */
    String format() default "";

    /**
     * 宽度
     */
    int width() default 0;

    /**
     * 顺序
     */
    int order() default 0;

    /**
     * 为null时的默认值
     */
    String defaultValue() default "";

    /**
     * 标记用于哪几个导出
     * <p>
     * 比如A字段只可以在管理员数据时导出，而在用户数据时不导出
     */
    int[] tags() default 0;

    /**
     * 处理拼接（可以和format共存，比如Double类型，可以先format为两位小数，后在拼接）
     * <p>
     * 例：第{{value}}月
     */
    String string() default "";

    /**
     * 分隔符
     * 在合并字段时使用
     */
    String separator() default "";
}

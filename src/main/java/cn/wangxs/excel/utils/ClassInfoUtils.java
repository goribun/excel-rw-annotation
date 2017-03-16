package cn.wangxs.excel.utils;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import cn.wangxs.excel.FieldInfo;
import cn.wangxs.excel.annotation.ExcelField;
import org.apache.commons.lang3.StringUtils;

/**
 * @author wangxuesong wangxuesong0302@autohome.com.cn
 * @version 1.0
 */
public class ClassInfoUtils {

    private ClassInfoUtils() {
    }

    /**
     * 字符串类型
     */
    public static final byte EXCEL_STRING_TYPE = 1;
    /**
     * 整数类型
     */
    public static final byte EXCEL_NUMBER_TYPE = 2;
    /**
     * 小数类型
     */
    public static final byte EXCEL_DECIMAL_TYPE = 3;
    /**
     * 日期类型
     */
    public static final byte EXCEL_DATE_TYPE = 4;

    /**
     * 替换的字符串
     */
    public static final String REPLACE_VALUE = "{{value}}";


    /**
     * 获取Object类型，根据类型常用情况，String类型先判断，然后是Integer等
     *
     * @param object 对象
     * @return excel对应的类型
     */
    public static byte getExcelTypeByObj(Object object) {

        if (object == null) {
            return EXCEL_STRING_TYPE;
        }

        String objString = object.getClass().toString();
        if (objString.contains("String")) {
            return EXCEL_STRING_TYPE;
        }

        if (isExcelNumberType(objString)) {
            return EXCEL_NUMBER_TYPE;
        }

        if (isExcelDecimalType(objString)) {
            return EXCEL_DECIMAL_TYPE;
        }

        if (objString.contains("Date")) {
            return EXCEL_DATE_TYPE;
        }

        //其余全部按字符串处理
        return EXCEL_STRING_TYPE;
    }

    //判断是否是EXCEL_NUMBER_TYPE
    private static boolean isExcelNumberType(String objString) {
        return objString.contains("Integer") || objString.contains("Long") || objString.contains("Short") || objString
                .contains("Byte") || objString.contains("BigInteger");
    }

    //判断是否是EXCEL_DECIMAL_TYPE
    private static boolean isExcelDecimalType(String objString) {
        return objString.contains("Double") || objString.contains("BigDecimal") || objString.contains("Float");
    }

    /**
     * 根据类返回field信息列表
     *
     * @param clazz 类
     * @return field信息列表
     * @throws IntrospectionException 内省异常
     */
    public static List<FieldInfo> getFieldInfoList(Class clazz) throws IntrospectionException {

        Field fields[] = clazz.getDeclaredFields();
        List<FieldInfo> fieldInfoList = new ArrayList<>();

        for (Field field : fields) {
            ExcelField exportField = field.getAnnotation(ExcelField.class);
            if (exportField == null) {
                continue;
            }
            //取得注解属性
            String name = StringUtils.isBlank(exportField.name()) ? field.getName() : exportField.name();
            String format = exportField.format();
            String defaultValue = exportField.defaultValue();
            String mergeTo = "";
            String separator = exportField.separator();
            String string = exportField.string();
            int[] tags = exportField.tags();
            int order = exportField.order();
            int width = exportField.width();
            //内省得到get信息
            PropertyDescriptor pd = new PropertyDescriptor(field.getName(), clazz);
            Method getMethod = pd.getReadMethod();
            //创建Field信息对象 TODO mergeTO后续完善
            FieldInfo fieldInfo = new FieldInfo(name, order, format, width, defaultValue, getMethod, mergeTo,
                    separator, string, tags);
            fieldInfoList.add(fieldInfo);
        }
        //设置排序
        sortList(fieldInfoList);
        return fieldInfoList;
    }

    /**
     * 根据order进行排序
     *
     * @param fieldInfoList field信息列表
     */
    private static void sortList(List<FieldInfo> fieldInfoList) {
        Collections.sort(fieldInfoList, new Comparator<FieldInfo>() {
            @Override
            public int compare(FieldInfo o1, FieldInfo o2) {
                if (o1.getOrder() > o2.getOrder()) {
                    return 1;
                } else if (o1.getOrder() < o2.getOrder()) {
                    return -1;
                }
                return 0;
            }
        });
    }
}
package cn.wangxs.excel.write.utils;

import java.beans.IntrospectionException;
import java.lang.reflect.InvocationTargetException;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import cn.wangxs.excel.FieldInfo;
import cn.wangxs.excel.utils.ClassInfoUtils;
import cn.wangxs.excel.write.row.RowContext;
import cn.wangxs.excel.write.sheet.SheetContext;
import cn.wangxs.excel.write.workbook.WorkbookContext;
import cn.wangxs.excel.write.workbook.WorkbookFactory;
import cn.wangxs.excel.write.workbook.WorkbookType;
import org.apache.commons.lang3.StringUtils;

/**
 * @author wangxuesong
 */
public class BaseWriteUtil {

    private BaseWriteUtil() {

    }

    /**
     * @param collection 集合
     * @param clazz      类
     * @param tag        标记（0标记会忽略）
     * @return excel
     */
    public static byte[] write(Collection<?> collection, Class<?> clazz, int tag) throws
            NoSuchFieldException,
            IllegalAccessException, IntrospectionException, InvocationTargetException {
        //创建workbook上下文
        WorkbookContext workbookContext = WorkbookFactory.createWorkbook(WorkbookType.SXSSF);
        List<FieldInfo> fieldInfoList = ClassInfoUtils.getFieldInfoList(clazz);
        //处理标记，0则忽略
        handleTags(tag, fieldInfoList);

        //取得excel表头
        List<String> headerList = new ArrayList<>();
        for (FieldInfo fieldInfo : fieldInfoList) {
            headerList.add(fieldInfo.getName());
        }

        //创建表头
        SheetContext sheetContext = workbookContext.createSheetAndHeader(headerList
                .toArray(new String[headerList.size()]));
        //内容空，直接返回
        if (collection == null || collection.size() == 0) {
            return workbookContext.toNativeBytes();
        }

        //根据不同类型写row和cell
        for (Object object : collection) {
            RowContext rowContext = sheetContext.nextRow();
            for (FieldInfo fieldInfo : fieldInfoList) {
                //值
                Object value = fieldInfo.getMethod().invoke(object);
                int excelType = ClassInfoUtils.getExcelTypeByObj(value);

                //处理需要特殊处理的字段，比如"第3周"
                String stringValue = handleSplicing(fieldInfo, value, excelType);
                if (stringValue != null) {
                    value = stringValue;
                    //类型指定处理为字符串
                    excelType = ClassInfoUtils.EXCEL_STRING_TYPE;
                }

                switch (excelType) {
                    case ClassInfoUtils.EXCEL_NUMBER_TYPE:
                        rowContext.number((Number) value);
                        break;
                    case ClassInfoUtils.EXCEL_DECIMAL_TYPE:
                        rowContext.decimal((Number) value, fieldInfo.getFormat());
                        break;
                    case ClassInfoUtils.EXCEL_DATE_TYPE:
                        rowContext.date((Date) value, fieldInfo.getFormat());
                        break;
                    default:
                        //处理字符串和其他类型，对象为null则处理为默认值
                        String obj = String.valueOf((value == null
                                || StringUtils.isBlank(value.toString())) ? fieldInfo.getDefaultValue() : value);
                        rowContext.text(obj);
                        break;
                }

                //宽度大于0 则设置宽度
                int width = fieldInfo.getWidth();
                if (width > 0) {
                    rowContext.setColumnWidth(width);
                }

            }
        }
        return workbookContext.toNativeBytes();
    }

    /**
     * 处理tags标记
     *
     * @param tag           标记
     * @param fieldInfoList field信息列表
     */
    private static void handleTags(int tag, List<FieldInfo> fieldInfoList) {
        //tag为0，表示忽略标记
        if (tag != 0) {
            Iterator<FieldInfo> iterator = fieldInfoList.iterator();
            while (iterator.hasNext()) {
                FieldInfo fieldInfo = iterator.next();
                if (fieldInfo.getTags().length == 1 && fieldInfo.getTags()[0] == 0) {
                    continue;
                }
                //如果注解tags里不包含参数tag，则remove掉（不导出）
                boolean contain = false;
                for (int tagTemp : fieldInfo.getTags()) {
                    if (tagTemp == tag) {
                        contain = true;
                    }
                }
                //删除
                if (!contain) {
                    iterator.remove();
                }
            }
        }
    }

    //处理需要特殊处理的字段
    private static String handleSplicing(FieldInfo fieldInfo, Object value, int excelType) {
        String string = fieldInfo.getString();
        String format = fieldInfo.getFormat();
        //没有处理需求或者值为null则不处理
        if (StringUtils.isBlank(string) || value == null) {
            return null;
        }

        if (StringUtils.isBlank(format)) {
            return string.replace(ClassInfoUtils.REPLACE_VALUE, value.toString());
        }
        if (excelType == ClassInfoUtils.EXCEL_DATE_TYPE) {
            SimpleDateFormat sdf = new SimpleDateFormat(format);
            return string.replace(ClassInfoUtils.REPLACE_VALUE, sdf.format(value));
        }
        if (excelType == ClassInfoUtils.EXCEL_DECIMAL_TYPE) {
            return NumberFormat.getInstance().format(value);
        }
        return null;
    }

}

package cn.wangxs.excel;

import java.beans.IntrospectionException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.Collection;
import java.util.List;
import java.util.Map;

import cn.wangxs.excel.read.utils.BaseReadUtil;
import cn.wangxs.excel.write.utils.BaseWriteUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

/**
 * 工具入口
 *
 * @author wangxuesong
 */
public class ExcelHelper {

    private ExcelHelper() {
    }

    /**
     * 生成excel文件的byte数组
     *
     * @param collection 数据集合
     * @param clazz      实体类型
     * @return byte数组
     */
    public static byte[] write(Collection<?> collection, Class<?> clazz) throws InvocationTargetException,
            IntrospectionException, IllegalAccessException, NoSuchFieldException {
        return BaseWriteUtil.write(collection, clazz, 0);
    }

    /**
     * 生成excel文件的byte数组
     * <p>
     * 只包含没有标记的字段以及标记匹配的字段
     *
     * @param collection 数据集合
     * @param clazz      实体类型
     * @param tag        标记
     * @return byte数组
     */
    public static byte[] write(Collection<?> collection, Class<?> clazz, int tag) throws InvocationTargetException,
            IntrospectionException, IllegalAccessException, NoSuchFieldException {
        return BaseWriteUtil.write(collection, clazz, tag);
    }

    /**
     * 读取excel文件为Map类型的列表
     *
     * @param is excel文件流
     * @return Map类型的列表
     */
    public static List<Map<String, String>> read(InputStream is) throws IOException, InvalidFormatException {
        return BaseReadUtil.read(is);
    }

    /**
     * 读取excel文件为泛型类型的列表
     *
     * @param is    excel文件流
     * @param clazz 实体类型
     * @param <T>   泛型类型
     * @return 泛型列表
     */
    public static <T> List<T> read(InputStream is, Class<T> clazz) throws IllegalAccessException,
            IntrospectionException, InvalidFormatException, IOException, InstantiationException,
            InvocationTargetException {

        return BaseReadUtil.read(is, clazz);
    }

}
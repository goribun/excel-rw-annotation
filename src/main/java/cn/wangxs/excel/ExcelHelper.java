package cn.wangxs.excel;

import java.io.InputStream;
import java.util.Collection;
import java.util.List;
import java.util.Map;

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
    public static byte[] write(Collection<?> collection, Class<?> clazz) {
        return null;
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
    public static byte[] write(Collection<?> collection, Class<?> clazz, int tag) {
        return null;
    }

    /**
     * 读取excel文件为Map类型的列表（）
     *
     * @param is excel文件流
     * @return Map类型的列表
     * @throws Exception 异常信息
     */
    public static List<Map<String, String>> read(InputStream is) throws Exception {
        return null;
    }

    /**
     * 读取excel文件为泛型类型的列表
     *
     * @param is    excel文件流
     * @param clazz 实体类型
     * @param <T>   泛型类型
     * @return 泛型列表
     * @throws Exception 异常信息
     */
    public static <T> List<T> read(InputStream is, Class<T> clazz) throws Exception {
        return null;
    }

}

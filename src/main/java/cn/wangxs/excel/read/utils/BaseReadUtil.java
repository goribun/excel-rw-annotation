package cn.wangxs.excel.read.utils;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import cn.wangxs.excel.FieldInfo;
import cn.wangxs.excel.annotation.ExcelField;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author wangxuesong
 * @version 1.0
 */
public class BaseReadUtil {
    /**
     * excel导入基础方法
     */
    public static <T> List<T> read(InputStream inputStream, Class<T> clazz) throws IOException,
            InvalidFormatException, IntrospectionException, InvocationTargetException, IllegalAccessException,
            InstantiationException {

        Sheet sheet = inputStream2Sheet(inputStream);
        if (sheet == null) {
            return Collections.emptyList();
        }

        List<T> list = new ArrayList<>();
        //字段信息列表
        Map<String, FieldInfo> fieldInfoMap = getExcelFieldList(clazz);

        //第一行为标题
        Iterator<Row> rowIterator = sheet.rowIterator();
        Row titleRow = rowIterator.next();
        //标题cell
        Iterator<Cell> cellTitleIterator = titleRow.cellIterator();

        List<String> titleList = new ArrayList<>();
        while (cellTitleIterator.hasNext()) {
            Cell cell = cellTitleIterator.next();
            String cellValue = cell.getStringCellValue();
            String title = StringUtils.isBlank(cellValue) ? "" : cellValue;
            titleList.add(title);
        }
        //循环处理数据行
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            if (isBlankRow(row)) {
                break;
            }

            T object = clazz.newInstance();
            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                Cell cell = row.getCell(i);
                if (cell == null) {
                    continue;
                }
                String titleString = titleList.get(i);
                //判断类型并赋值
                judgeTypeAndSetValue(object, cell, fieldInfoMap, titleString);
            }
            list.add(object);
        }
        return list;
    }

    /**
     * 判断类型并设置值
     */
    private static void judgeTypeAndSetValue(Object object, Cell cell, Map<String, FieldInfo> fieldInfoMap,
                                             String titleString) throws InvocationTargetException,
            IllegalAccessException {
        FieldInfo entity = fieldInfoMap.get(titleString);

        if (entity == null) {
            return;
        }

        Method setMethod = entity.getMethod();
        Type[] ts = setMethod.getGenericParameterTypes();
        String fieldClass = ts[0].toString();
        switch (fieldClass) {
            case "class java.lang.String":
                cell.setCellType(Cell.CELL_TYPE_STRING);
                String cellValue = cell.getStringCellValue();
                if (StringUtils.isNotEmpty(cellValue)) {
                    cellValue = cellValue.trim();
                }
                entity.getMethod().invoke(object, cellValue);
                break;
            case "class java.util.Date":
                Date cellDate;
                if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                    // 日期格式
                    cellDate = cell.getDateCellValue();
                    entity.getMethod().invoke(object, cellDate);
                } else {
                    cell.setCellType(Cell.CELL_TYPE_STRING);
                    cellDate = getDateData(entity, cell.getStringCellValue());
                    entity.getMethod().invoke(object, cellDate);
                }
                break;
            case "class java.lang.Boolean":
                boolean valBool;
                if (Cell.CELL_TYPE_BOOLEAN == cell.getCellType()) {
                    valBool = cell.getBooleanCellValue();
                } else {
                    valBool = cell.getStringCellValue().equalsIgnoreCase("true")
                            || (!cell.getStringCellValue().equals("0"));
                }
                entity.getMethod().invoke(object, valBool);

                break;
            case "class java.lang.Integer":
                Integer valInt;
                if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                    valInt = Double.valueOf(cell.getNumericCellValue()).intValue();
                } else {
                    valInt = Integer.valueOf(cell.getStringCellValue());
                }
                entity.getMethod().invoke(object, valInt);

                break;
            case "class java.lang.Long":
                Long valLong;
                if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                    valLong = Double.valueOf(cell.getNumericCellValue()).longValue();
                } else {
                    valLong = Long.valueOf(cell.getStringCellValue());
                }
                entity.getMethod().invoke(object, valLong);

                break;
            case "class java.math.BigDecimal":
                BigDecimal valDecimal;
                if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                    valDecimal = new BigDecimal(cell.getNumericCellValue());
                } else {
                    valDecimal = new BigDecimal(cell.getStringCellValue());
                }
                entity.getMethod().invoke(object, valDecimal);
                break;
            case "class java.lang.Double":
                Double valDouble;
                if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                    valDouble = cell.getNumericCellValue();
                } else {
                    valDouble = new Double(cell.getStringCellValue());
                }
                entity.getMethod().invoke(object, valDouble);
                break;
            default:
                cell.setCellType(Cell.CELL_TYPE_STRING);
                String cellDefaultValue = cell.getStringCellValue();
                if (StringUtils.isNotEmpty(cellDefaultValue)) {
                    cellDefaultValue = cellDefaultValue.trim();
                }
                entity.getMethod().invoke(object, cellDefaultValue);
                break;
        }
    }

    private static Map<String, FieldInfo> getExcelFieldList(Class<?> clazz) throws IntrospectionException {
        Field[] fields = clazz.getDeclaredFields();
        Map<String, FieldInfo> map = new HashMap<>();
        for (Field field : fields) {
            ExcelField importField = field.getAnnotation(ExcelField.class);
            if (importField == null) {
                continue;
            }
            FieldInfo fieldInfo = new FieldInfo();
            getExcelField(field, fieldInfo, importField, clazz);
            map.put(fieldInfo.getName(), fieldInfo);
        }
        return map;
    }

    /**
     * 获取导入字段
     */
    private static void getExcelField(Field field, FieldInfo fieldInfo, ExcelField excelField, Class<?> clazz) throws
            IntrospectionException {

        fieldInfo.setName(excelField.name());
        PropertyDescriptor pd = new PropertyDescriptor(field.getName(), clazz);
        Method setMethod = pd.getWriteMethod();
        fieldInfo.setMethod(setMethod);

        fieldInfo.setFormat(excelField.format());
    }

    /**
     * 获取日期类型数据
     */
    private static Date getDateData(FieldInfo entity, String value) {
        if (StringUtils.isNotEmpty(entity.getFormat())
                && StringUtils.isNotEmpty(value)) {
            SimpleDateFormat format = new SimpleDateFormat(entity.getFormat());
            try {
                return format.parse(value);
            } catch (ParseException e) {
                return null;
            }
        }
        return null;
    }

    /**
     * 判断空行
     */
    private static boolean isBlankRow(Row row) {
        if (row == null) {
            return true;
        }
        boolean result = true;
        Iterator<Cell> cells = row.cellIterator();
        String value = "";
        while (cells.hasNext()) {
            Cell cell = cells.next();
            int cellType = cell.getCellType();
            switch (cellType) {
                case Cell.CELL_TYPE_NUMERIC:
                    value = String.valueOf(cell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    value = String.valueOf(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    value = String.valueOf(cell.getCellFormula());
                    break;
            }
            if (StringUtils.isNotBlank(value)) {
                result = false;
                break;
            }
        }

        return result;
    }

    /**
     * 处理excel到List(用于简单导入，不返回具体实体的列表／简单导入使用)
     *
     * @param inputStream 文件流
     * @return 数据列表
     */
    public static List<Map<String, String>> read(InputStream inputStream) throws IOException,
            InvalidFormatException {


        Sheet sheet = inputStream2Sheet(inputStream);
        if (sheet == null) {
            return Collections.emptyList();
        }

        Row row;
        Iterator<Row> rows = sheet.rowIterator();
        //表头
        List<String> titleList = new ArrayList<>();

        List<Map<String, String>> result = new ArrayList<>();
        //是否是第一行
        boolean firstRow = true;
        while (rows.hasNext()) {
            //每行数据的Map，key为表头，value为值
            Map<String, String> map = new HashMap<>();

            row = rows.next();
            if (isBlankRow(row)) {
                break;
            }

            for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
                Cell cell = row.getCell(i);
                if (cell == null) {
                    continue;
                }
                //取得cell值
                String cellValue = getCellFormatValue(cell);
                if (firstRow) {
                    titleList.add(cellValue.trim());
                } else {
                    if (i < titleList.size()) {
                        map.put(titleList.get(i), cellValue.trim());
                    }
                }
            }
            //每行记录加入
            if (map.size() > 0) {
                result.add(map);
            }
            //第一行设置为否
            firstRow = false;
        }
        return result;
    }


    /**
     * 根据HSSFCell类型设置数据（简单导入使用）
     *
     * @param cell 单元格
     * @return 处理后的值
     */
    private static String getCellFormatValue(Cell cell) {
        String cellValue;
        if (cell != null) {
            // 判断当前Cell的Type
            switch (cell.getCellType()) {
                // 如果当前Cell的Type为NUMERIC
                case Cell.CELL_TYPE_NUMERIC:
                case Cell.CELL_TYPE_FORMULA: {
                    // 判断当前的cell是否为Date
                    if (DateUtil.isCellDateFormatted(cell)) {
                        // 如果是Date类型则，转化为Data格式
                        Date date = cell.getDateCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                        cellValue = sdf.format(date);
                    }
                    // 如果是纯数字
                    else {
                        //避免科学计数
                        java.text.NumberFormat nf = java.text.NumberFormat.getInstance();
                        nf.setGroupingUsed(false);
                        // 取得当前Cell的数值
                        cellValue = String.valueOf(nf.format(cell.getNumericCellValue()));
                    }
                    break;
                }
                // 如果当前Cell的Type为STRING
                case Cell.CELL_TYPE_STRING:
                    // 取得当前的Cell字符串
                    cellValue = cell.getStringCellValue();
                    break;
                // 默认的Cell值
                default:
                    cellValue = " ";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }

    private static Sheet inputStream2Sheet(InputStream inputStream) throws IOException, InvalidFormatException {

        Workbook book = null;
        if (!inputStream.markSupported()) {
            inputStream = new PushbackInputStream(inputStream, 8);
        }
        if (POIFSFileSystem.hasPOIFSHeader(inputStream)) {
            book = new HSSFWorkbook(inputStream);
        } else if (POIXMLDocument.hasOOXMLHeader(inputStream)) {
            book = new XSSFWorkbook(OPCPackage.open(inputStream));
        }
        if (book == null) {
            return null;
        }
        //默认取得第一个sheet
        return book.getSheetAt(0);
    }
}
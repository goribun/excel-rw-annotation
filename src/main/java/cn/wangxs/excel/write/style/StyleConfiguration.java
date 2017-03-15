package cn.wangxs.excel.write.style;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 样式配置类
 *
 * @author wangxuesong
 */

public class StyleConfiguration {

    private Workbook workbook;//workbook


    private static final Byte DEFAULT_STYLE_KEY = 10;
    private static final Byte HEADER_STYLE_KEY = 11;
    private static final Byte DECIMAL_STYLE_KEY = 12;
    private static final Byte DATE_STYLE_KEY = 13;
    private static final Byte DATE_8_STYLE_KEY = 14;

    private static final Byte DATA_FORMAT_KEY = 20;

    private static final Byte FONT_KEY = 30;

    private final Map<Byte, CellStyle> buildInStyleMap = new HashMap<>(8);//内建样式
    private final Map<Byte, DataFormat> buildInFormatMap = new HashMap<>(2);//内建格式化
    private final Map<Byte, Font> buildInFontMap = new HashMap<>(2);//内建字体

    private final Map<String, CellStyle> customFormatStyleMap = new HashMap<>(8);//用户自定义格式化样式

    public StyleConfiguration(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * header样式
     *
     * @return CellStyle
     */
    public CellStyle getHeaderStyle() {

        if (buildInStyleMap.containsKey(HEADER_STYLE_KEY)) {
            return buildInStyleMap.get(HEADER_STYLE_KEY);
        }

        CellStyle headerStyle = workbook.createCellStyle();//头的样式
        // 设置单元格的背景颜色为淡蓝色
        headerStyle.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
        headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        this.setCommonStyle(headerStyle);
        buildInStyleMap.put(HEADER_STYLE_KEY, headerStyle);
        return headerStyle;
    }

    /**
     * 文本样式
     *
     * @return CellStyle
     */
    public CellStyle getTextStyle() {
        //text使用默认style
        return this.getDefaultStyle();
    }


    /**
     * 数字样式
     *
     * @return CellStyle
     */
    public CellStyle getNumberStyle() {
        //number使用默认style
        return this.getDefaultStyle();
    }


    /**
     * 小数格式
     *
     * @return CellStyle
     */
    public CellStyle getDecimalStyle() {
        if (buildInStyleMap.containsKey(DECIMAL_STYLE_KEY)) {
            return buildInStyleMap.get(DECIMAL_STYLE_KEY);
        }

        CellStyle decimalStyle = workbook.createCellStyle();//小数样式
        if (!buildInFormatMap.containsKey(DATA_FORMAT_KEY)) {
            DataFormat dataFormat = workbook.createDataFormat();
            buildInFormatMap.put(DATA_FORMAT_KEY, dataFormat);
        }
        decimalStyle.setDataFormat(buildInFormatMap.get(DATA_FORMAT_KEY).getFormat("0.00"));
        this.setCommonStyle(decimalStyle);
        buildInStyleMap.put(DECIMAL_STYLE_KEY, decimalStyle);
        return decimalStyle;
    }

    /**
     * 日期样式 yyyy-MM-dd HH:mm
     *
     * @return CellStyle
     */
    public CellStyle getDateStyle() {
        if (buildInStyleMap.containsKey(DATE_STYLE_KEY)) {
            return buildInStyleMap.get(DATE_STYLE_KEY);
        }

        CellStyle dateStyle = workbook.createCellStyle();//日期样式
        if (!buildInFormatMap.containsKey(DATA_FORMAT_KEY)) {
            DataFormat dataFormat = workbook.createDataFormat();
            buildInFormatMap.put(DATA_FORMAT_KEY, dataFormat);
        }
        dateStyle.setDataFormat(buildInFormatMap.get(DATA_FORMAT_KEY).getFormat("yyyy-MM-dd HH:mm"));
        this.setCommonStyle(dateStyle);
        buildInStyleMap.put(DATE_STYLE_KEY, dateStyle);
        return dateStyle;
    }

    /**
     * 日期样式 yyyy/MM/dd
     *
     * @return CellStyle
     */
    public CellStyle getDate8Style() {

        if (buildInStyleMap.containsKey(DATE_8_STYLE_KEY)) {
            return buildInStyleMap.get(DATE_8_STYLE_KEY);
        }

        CellStyle date8Style = workbook.createCellStyle();//年月日样式
        if (!buildInFormatMap.containsKey(DATA_FORMAT_KEY)) {
            DataFormat dataFormat = workbook.createDataFormat();
            buildInFormatMap.put(DATA_FORMAT_KEY, dataFormat);
        }
        date8Style.setDataFormat(buildInFormatMap.get(DATA_FORMAT_KEY).getFormat("yyyy/MM/dd"));
        this.setCommonStyle(date8Style);
        buildInStyleMap.put(DATE_8_STYLE_KEY, date8Style);
        return date8Style;
    }

    /**
     * 根据格式，创建返回样式对象
     *
     * @param format 格式
     * @return 样式对象
     */
    public CellStyle getCustomFormatStyle(String format) {

        //存在对应格式直接返回
        if (customFormatStyleMap.containsKey(format)) {
            return customFormatStyleMap.get(format);
        }
        CellStyle customDateStyle = workbook.createCellStyle();
        if (!buildInFormatMap.containsKey(DATA_FORMAT_KEY)) {
            DataFormat dataFormat = workbook.createDataFormat();
            buildInFormatMap.put(DATA_FORMAT_KEY, dataFormat);
        }
        customDateStyle.setDataFormat(buildInFormatMap.get(DATA_FORMAT_KEY).getFormat(format));
        this.setCommonStyle(customDateStyle);
        //放入map缓存
        customFormatStyleMap.put(format, customDateStyle);

        return customDateStyle;
    }

    /**
     * 默认样式,目前文字和整数使用的该样式
     *
     * @return CellStyle
     */
    public CellStyle getDefaultStyle() {

        if (buildInStyleMap.containsKey(DEFAULT_STYLE_KEY)) {
            return buildInStyleMap.get(DEFAULT_STYLE_KEY);
        }

        CellStyle defaultStyle = workbook.createCellStyle();//默认样式
        // 设置单元格边框为细线条
        this.setCommonStyle(defaultStyle);
        buildInStyleMap.put(DEFAULT_STYLE_KEY, defaultStyle);
        return defaultStyle;
    }

    /**
     * 设置通用的对齐居中、边框等
     *
     * @param style 样式
     */
    private void setCommonStyle(CellStyle style) {
        // 设置单元格居中对齐、自动换行
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setWrapText(true);

        //设置单元格字体
        if (!buildInFontMap.containsKey(FONT_KEY)) {
            Font font = workbook.createFont();
            //通用字体
            font.setBoldweight(Font.BOLDWEIGHT_BOLD);
            font.setFontName("宋体");
            font.setFontHeight((short) 200);
            buildInFontMap.put(FONT_KEY, font);
        }
        style.setFont(buildInFontMap.get(FONT_KEY));

        // 设置单元格边框为细线条
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setBorderTop(CellStyle.BORDER_THIN);
    }
}
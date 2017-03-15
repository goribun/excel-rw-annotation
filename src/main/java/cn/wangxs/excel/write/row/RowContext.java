package cn.wangxs.excel.write.row;

import java.util.Date;

import cn.wangxs.excel.write.sheet.SheetContext;
import cn.wangxs.excel.write.style.StyleConfiguration;
import cn.wangxs.excel.write.workbook.WorkbookContext;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

/**
 * 原生Row的包装类
 *
 * @author wangxuesong
 */
public class RowContext {

    private static final int ROW_HEIGHT_AUTOMATIC = -1;
    private final Row row;
    private final SheetContext sheet;
    private int index;
    private final short rowHeight;
    private int step;
    private final StyleConfiguration styleConfiguration;

    public RowContext(Row row, SheetContext sheet, WorkbookContext workbookContext) {
        this.row = row;
        this.sheet = sheet;
        this.index = 0;
        this.step = 1;
        this.rowHeight = ROW_HEIGHT_AUTOMATIC;
        styleConfiguration = workbookContext.getStyleConfiguration();
    }

    /**
     * 写字符串方法
     *
     * @param text 字符串
     */
    public RowContext text(String text) {
        return writeText(text, styleConfiguration.getTextStyle());
    }

    /**
     * 写字符串方法
     * 重载方法，支持自定义CellStyle
     *
     * @param text  字符串
     * @param style 样式
     */
    public RowContext text(String text, CellStyle style) {
        return writeText(text, style);

    }

    /**
     * 写整数
     *
     * @param number 整数
     */
    public RowContext number(Number number) {
        return writeNumber(number, styleConfiguration.getNumberStyle());
    }

    /**
     * 写整数
     * 重载的方法，支持自定义CellStyle
     *
     * @param number 整数
     * @param style  样式
     */
    public RowContext number(Number number, CellStyle style) {
        return writeNumber(number, style);
    }

    /**
     * 写小数
     *
     * @param number 小数
     */
    public RowContext decimal(Number number) {
        return writeNumber(number, styleConfiguration.getDecimalStyle());
    }

    /**
     * 写小数
     * 重载的方法，支持自定义格式化
     *
     * @param number 小数
     * @param format 格式化
     */
    public RowContext decimal(Number number, String format) {
        return writeNumber(number, styleConfiguration.getCustomFormatStyle(format));
    }


    /**
     * 写日期，默认格式yyyy-MM-dd HH:mm
     *
     * @param date 日期
     */
    public RowContext date(Date date) {
        return writeDate(date, styleConfiguration.getDateStyle());
    }

    /**
     * 写日期
     * 重载的方法，支持自定义格式化
     *
     * @param date   日期
     * @param format 格式化格式
     */
    public RowContext date(Date date, String format) {
        return writeDate(date, styleConfiguration.getCustomFormatStyle(format));
    }

    /**
     * 写日期
     * 重载的方法，支持自定义CellStyle
     *
     * @param date  日期
     * @param style 样式
     */
    public RowContext date(Date date, CellStyle style) {
        return writeDate(date, style);
    }

    /**
     * 写日期，默认格式yyyy/MM/dd
     *
     * @param date 日期
     */
    public RowContext date8(Date date) {
        return writeDate(date, styleConfiguration.getDate8Style());
    }

    /**
     * 写表头
     *
     * @param header 表头字符串
     */
    public RowContext header(String header) {
        return writeText(header, styleConfiguration.getHeaderStyle());
    }

    /**
     * 写表头
     * 重载方法 支持自定义CellStyle
     *
     * @param header 表头字符串
     * @param style  样式
     */
    public RowContext header(String header, CellStyle style) {
        return writeText(header, style);
    }

    public RowContext setColumnWidth(int width) {
        sheet.setColumnWidth(index - 1, width);
        return this;
    }


    /**
     * 跳过一个cell
     */
    public RowContext skipOneCell() {
        return skipCells(1);
    }

    /**
     * 跳过n个cell
     *
     * @param offset 偏移量
     */
    public RowContext skipCells(int offset) {
        index += offset;
        return this;
    }


    /**
     * 写text
     * 2016/12/09增加判空
     *
     * @param text  文本
     * @param style 样式
     */
    private RowContext writeText(String text, CellStyle style) {
        createCell(1, style).setCellValue(StringUtils.isBlank(text) ? "" : text);
        return this;

    }


    /**
     * 写number
     *
     * @param number 数字
     * @param style  样式
     */
    private RowContext writeNumber(Number number, CellStyle style) {
        if (number == null) {
            number = 0;
        }
        createCell(1, style).setCellValue(number.doubleValue());
        return this;
    }

    /**
     * 写date
     * 2016/12/09增加判空
     *
     * @param date  日期
     * @param style 样式
     */
    private RowContext writeDate(Date date, CellStyle style) {

        if (date == null) {
            writeText("", style);
            return this;
        }
        //修正15:27:59.583中的583毫秒进位问题(将毫秒至为000)
        Date result = DateUtils.setMilliseconds(date, 0);
        createCell(1, style).setCellValue(result);
        return this;
    }

    /**
     * 创建cell
     */
    private Cell createCell(int rowHeightMultiplier, CellStyle style) {
        assignRowHeight(rowHeightMultiplier);
        Cell cell = row.createCell(index);
        cell.setCellStyle(style);
        index += step;
        step = 1;
        return cell;
    }

    private void assignRowHeight(int rowHeightMultiplier) {
        if (rowHeightMultiplier > 1 && rowHeight == ROW_HEIGHT_AUTOMATIC) {
            row.setHeightInPoints(row.getHeightInPoints() * rowHeightMultiplier);
        } else {
            row.setHeight(rowHeight);
        }
    }
}
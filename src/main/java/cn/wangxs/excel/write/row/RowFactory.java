package cn.wangxs.excel.write.row;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Row工厂
 * Created by wangxuesong on 16/3/11.
 */
public final class RowFactory {

    private RowFactory() {
    }

    public static Row getOrCreate(Sheet sheet, int index) {
        if (sheet == null) {
            throw new NullPointerException("sheet is null !");
        }
        Row row = sheet.getRow(index);
        if (row == null) {
            return sheet.createRow(index);
        }
        return row;
    }
}

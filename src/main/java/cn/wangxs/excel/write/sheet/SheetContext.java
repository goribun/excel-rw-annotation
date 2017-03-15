package cn.wangxs.excel.write.sheet;

import cn.wangxs.excel.write.row.RowContext;
import cn.wangxs.excel.write.row.RowFactory;
import cn.wangxs.excel.write.workbook.WorkbookContext;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * 原生Sheet包装类
 *
 * @author wangxuesong
 */
public class SheetContext {

    private final WorkbookContext workbookContext;
    private final Sheet sheet;
    private RowContext currentRow;
    private int index = -1;

    public SheetContext(WorkbookContext workbookContext, Sheet sheet) {
        this.workbookContext = workbookContext;
        this.sheet = sheet;
    }


    public RowContext nextRow() {
        ++index;
        currentRow = null;
        return getCurrentRow();
    }

    public RowContext getCurrentRow() {
        if (index == -1) {
            return null;
        }
        if (currentRow == null) {
            currentRow = new RowContext(RowFactory.getOrCreate(sheet, index), this, workbookContext);
        }
        return currentRow;
    }

    public SheetContext setColumnWidths(int... columnNo) {
        for (int i = 0; i < columnNo.length; i++) {
            setColumnWidth(i, columnNo[i]);
        }
        return this;
    }

    public SheetContext hideGrid() {
        sheet.setDisplayGridlines(true);
        return this;
    }

    public SheetContext setColumnWidth(int columnNo, int width) {
        sheet.setColumnWidth(columnNo, width);
        return this;
    }

    public SheetContext skipOneRow() {
        return skipRows(1);
    }

    public SheetContext skipRows(int offset) {
        this.index += offset;
        return this;
    }
}
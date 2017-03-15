package cn.wangxs.excel.write.workbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

import cn.wangxs.excel.write.row.RowContext;
import cn.wangxs.excel.write.sheet.SheetContext;
import cn.wangxs.excel.write.style.StyleConfiguration;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 原生Workbook包装类
 *
 * @author wangxuesong
 */
public class WorkbookContext {

    private final StyleConfiguration styleConfiguration;
    private final Workbook workbook;

    WorkbookContext(Workbook workbook) {
        this.workbook = workbook;
        this.styleConfiguration = new StyleConfiguration(workbook);
    }

    /**
     * 创建Sheet
     *
     * @param sheetName sheet名
     */
    public SheetContext createSheet(String sheetName) {
        return new SheetContext(this, workbook.createSheet(sheetName));
    }

    /**
     * 创建名称为空的Sheet
     */
    public SheetContext createSheet() {
        return new SheetContext(this, workbook.createSheet());
    }


    /**
     * 创建名称为空的Sheet并生成header行
     *
     * @param headerArray 表头数组
     */
    public SheetContext createSheetAndHeader(String... headerArray) {

        SheetContext sheetContext = createSheet();
        RowContext headerContext = sheetContext.nextRow();
        for (int i = 0; i < headerArray.length; i++) {
            headerContext.header(headerArray[i]);
        }

        return sheetContext;
    }

    /**
     * 原生Bytes
     */
    public byte[] toNativeBytes() {
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            workbook.write(baos);
            return baos.toByteArray();
        } catch (IOException e) {
            throw new RuntimeException("ToNativeBytes Failed ", e);
        }
    }

    public StyleConfiguration getStyleConfiguration() {
        return this.styleConfiguration;
    }

    /**
     * 原生WorkBook
     */
    public Workbook toNativeWorkbook() {
        return workbook;
    }
}
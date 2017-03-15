package cn.wangxs.excel.write.workbook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Workbook工厂
 *
 * @author wangxuesong
 */
public final class WorkbookFactory {

    private WorkbookFactory() {
    }

    /**
     * 创建Workbook
     *
     * @return WorkbookContext
     */
    public static WorkbookContext createWorkbook() {
        Workbook workbook = new HSSFWorkbook();
        return new WorkbookContext(workbook);
    }

    public static WorkbookContext createWorkbook(WorkbookType workbookType) {

        //创建XLSX格式的excel
        if (workbookType == WorkbookType.XSSF) {
            Workbook workbook = new XSSFWorkbook();
            return new WorkbookContext(workbook);
        }

        if (workbookType == WorkbookType.SXSSF) {
            Workbook workbook = new SXSSFWorkbook();
            return new WorkbookContext(workbook);
        }
        //默认返回的Workbook
        return createWorkbook();
    }
}
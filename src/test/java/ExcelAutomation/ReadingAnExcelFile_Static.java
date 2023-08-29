package ExcelAutomation;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;

public class ReadingAnExcelFile_Static {

    @Test
    public void xls_Old() throws Exception {
        File file = new File("ExcelFiles\\Excel.xls");
        FileInputStream fis = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(fis);
        HSSFSheet sheet = workbook.getSheet("Test");

        HSSFRow row = sheet.getRow(0);
        HSSFCell cell = row.getCell(0);
        String value = cell.getStringCellValue();
        System.out.println(value);

        workbook.close();
        fis.close();
    }

    @Test
    public void xlsx_New() throws Exception {
        File file = new File("ExcelFiles\\Excel.xlsx");
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Test");

        XSSFRow row = sheet.getRow(0);
        XSSFCell cell = row.getCell(1);
        String value = cell.getStringCellValue();
        System.out.println(value);

        workbook.close();
        fis.close();
    }
}
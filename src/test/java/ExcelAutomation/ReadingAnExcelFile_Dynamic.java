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

public class ReadingAnExcelFile_Dynamic {

    @Test
    public void xls_Old() throws Exception {
        File file = new File("ExcelFiles\\Excel.xls");
        FileInputStream fis = new FileInputStream(file);
        HSSFWorkbook workbook = new HSSFWorkbook(fis);
        HSSFSheet sheet = workbook.getSheet("Test");

        int rowCount = sheet.getPhysicalNumberOfRows();

        for (int i = 0; i < rowCount; i++) {
            HSSFRow row = sheet.getRow(i);
            int columnCount = row.getPhysicalNumberOfCells();

            for (int j = 0; j < columnCount; j++) {
                HSSFCell cell = row.getCell(j);
                String cellValue = cell.getStringCellValue();
                System.out.println(cellValue);
            }
        }

        workbook.close();
        fis.close();
    }

    @Test
    public void xlsx_New() throws Exception {
        File file = new File("ExcelFiles\\Excel.xlsx");
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Test");

        int rowCount = sheet.getPhysicalNumberOfRows();

        for (int i = 0; i < rowCount; i++) {
            XSSFRow row = sheet.getRow(i);
            int columnCount = row.getPhysicalNumberOfCells();

            for (int j = 0; j < columnCount; j++) {
                XSSFCell cell = row.getCell(j);
                String cellValue = cell.getStringCellValue();
                System.out.println(cellValue);
            }
        }

        workbook.close();
        fis.close();
    }
}
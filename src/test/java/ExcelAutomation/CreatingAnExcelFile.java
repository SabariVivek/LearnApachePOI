package ExcelAutomation;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;

public class CreatingAnExcelFile {

    @Test
    public void xls_Old() throws Exception {
        deleteTheExistingFile("xlx");

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Test");

        sheet.createRow(0);
        sheet.getRow(0).createCell(0).setCellValue("Username");
        sheet.getRow(0).createCell(1).setCellValue("Password");

        sheet.createRow(1);
        sheet.getRow(1).createCell(0).setCellValue("Sabari");
        sheet.getRow(1).createCell(1).setCellValue("WhoAmI@123");

        workbook.write(new File("ExcelFiles\\Excel.xls"));
        workbook.close();

        File file = new File("ExcelFiles\\Excel.xls");
        Desktop.getDesktop().browse(file.toURI());
    }

    @Test
    public void xlsx_New() throws Exception {
        deleteTheExistingFile("xlsx");

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Test");

        sheet.createRow(0);
        sheet.getRow(0).createCell(0).setCellValue("Username");
        sheet.getRow(0).createCell(1).setCellValue("Password");

        sheet.createRow(1);
        sheet.getRow(1).createCell(0).setCellValue("Sabari");
        sheet.getRow(1).createCell(1).setCellValue("WhoAmI@123");

        File file = new File("ExcelFiles\\Excel.xlsx");
        FileOutputStream fos = new FileOutputStream(file);
        workbook.write(fos);
        workbook.close();

        File fileOpen = new File("ExcelFiles\\Excel.xlsx");
        Desktop.getDesktop().browse(file.toURI());
    }

    public void deleteTheExistingFile(String extension) {
        try {
            File file = new File("ExcelFiles\\Excel." + extension.trim());
            if (file.exists()) {
                file.delete();
            }
        } catch (Exception e) {
            System.out.println("Deleting the existing file step failed");
        }
    }
}
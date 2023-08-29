package ExcelAutomation;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

public class ReadingExcelDataAsMap {

    @Test
    public void gettingExcelDataUsingMap() throws Exception {
        Map<String, String> map = new HashMap<>();

        File file = new File("ExcelFiles\\Excel.xlsx");
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheet("Test");

        int rowCount = sheet.getPhysicalNumberOfRows();

        for (int i = 1; i < rowCount; i++) {
            XSSFRow row = sheet.getRow(i);
            int columnCount = row.getPhysicalNumberOfCells();

            for (int j = 0; j < columnCount; j++) {
                XSSFCell cell = row.getCell(j);
                String cellValue = cell.getStringCellValue();
                map.put(sheet.getRow(0).getCell(j).getStringCellValue(), cellValue);
            }
            System.out.println(map);
        }

        System.out.println();
        workbook.close();
        fis.close();
    }
}
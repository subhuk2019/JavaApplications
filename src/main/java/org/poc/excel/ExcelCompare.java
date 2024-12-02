package org.poc.excel;

//TIP To <b>Run</b> code, press <shortcut actionId="Run"/> or
// click the <icon src="AllIcons.Actions.Execute"/> icon in the gutter.
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

public class ExcelCompare {

    public static void main(String[] args) {
        File file1 = new File("path_to_first_excel_file.xlsx");
        File file2 = new File("path_to_second_excel_file.xlsx");

        try {
            // Load both workbooks
            FileInputStream fis1 = new FileInputStream(file1);
            FileInputStream fis2 = new FileInputStream(file2);
            Workbook wb1 = new XSSFWorkbook(fis1);
            Workbook wb2 = new XSSFWorkbook(fis2);

            // Compare the workbooks
            compareExcelFiles(wb1, wb2);

            // Close resources
            fis1.close();
            fis2.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void compareExcelFiles(Workbook wb1, Workbook wb2) {
        // Check if both workbooks have the same number of sheets
        if (wb1.getNumberOfSheets() != wb2.getNumberOfSheets()) {
            System.out.println("The Excel files have a different number of sheets.");
            return;
        }

        // Compare each sheet
        for (int i = 0; i < wb1.getNumberOfSheets(); i++) {
            Sheet sheet1 = wb1.getSheetAt(i);
            Sheet sheet2 = wb2.getSheetAt(i);

            // Compare the number of rows in the current sheet
            if (sheet1.getPhysicalNumberOfRows() != sheet2.getPhysicalNumberOfRows()) {
                System.out.println("Sheets " + sheet1.getSheetName() + " have different row counts.");
                return;
            }

            // Compare each row in the sheet
            for (int rowIndex = 0; rowIndex < sheet1.getPhysicalNumberOfRows(); rowIndex++) {
                Row row1 = sheet1.getRow(rowIndex);
                Row row2 = sheet2.getRow(rowIndex);

                // Compare each cell in the row
                if (row1 != null && row2 != null) {
                    for (int colIndex = 0; colIndex < row1.getPhysicalNumberOfCells(); colIndex++) {
                        Cell cell1 = row1.getCell(colIndex);
                        Cell cell2 = row2.getCell(colIndex);

                        if (cell1 != null && cell2 != null) {
                            String value1 = cell1.toString();
                            String value2 = cell2.toString();

                            if (!value1.equals(value2)) {
                                System.out.println("Difference found at sheet: " + sheet1.getSheetName() +
                                        ", row: " + (rowIndex + 1) + ", column: " + (colIndex + 1) +
                                        " | " + value1 + " != " + value2);
                            }
                        }
                    }
                }
            }
        }
    }
}

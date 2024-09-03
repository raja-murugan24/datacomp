import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class ExcelExceptionUpdaterOptimized {

    public static void main(String[] args) {
        String filePath = "path/to/your/excel/file.xlsx"; // Update with your Excel file path
        
        try (FileInputStream fileInputStream = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Read the "Differences_Data" sheet
            Sheet differencesSheet = workbook.getSheet("Differences_Data");
            // Read the "Exception" sheet
            Sheet exceptionSheet = workbook.getSheet("Exception");

            // Store Employee_IDs from "Exception" sheet in a Map for fast lookup
            Map<String, Boolean> exceptionEmployeeIds = new HashMap<>();
            for (int i = 1; i <= exceptionSheet.getLastRowNum(); i++) { // Start from 1 to skip header
                Row row = exceptionSheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(0); // Employee_ID column
                    if (cell != null) {
                        exceptionEmployeeIds.put(cell.getStringCellValue(), true);
                    }
                }
            }

            // Buffer for collecting rows that need updating
            boolean needsUpdate = false;

            // Iterate through the "Differences_Data" sheet and update the note if Employee_ID is in exception list
            for (int i = 1; i <= differencesSheet.getLastRowNum(); i++) { // Start from 1 to skip header
                Row row = differencesSheet.getRow(i);
                if (row != null) {
                    Cell employeeIdCell = row.getCell(0); // Employee_ID column in Differences_Data sheet
                    if (employeeIdCell != null && exceptionEmployeeIds.containsKey(employeeIdCell.getStringCellValue())) {
                        // Add the note in the last column
                        int lastColumnIndex = row.getLastCellNum();
                        Cell noteCell = row.createCell(lastColumnIndex);
                        noteCell.setCellValue("This Employee is Exception. Ignore him");
                        needsUpdate = true;
                    }
                }
            }

            // Write the changes to the Excel file only if there is an update needed
            if (needsUpdate) {
                try (FileOutputStream fileOutputStream = new FileOutputStream(filePath)) {
                    workbook.write(fileOutputStream);
                }
                System.out.println("Excel file updated successfully!");
            } else {
                System.out.println("No updates were needed.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

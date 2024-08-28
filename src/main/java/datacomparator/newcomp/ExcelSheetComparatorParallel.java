package datacomparator.newcomp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class ExcelSheetComparatorParallel {

    public static void main(String[] args) {
        String filePath = "your-excel-file-path.xlsx"; // Update with your file path
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sitSheet = workbook.getSheet("SIT");
            Sheet prodSheet = workbook.getSheet("PROD");

            // Create the Differences sheet as the 3rd sheet, delete if already exists
            int differencesSheetIndex = 2; // 0-based index
            Sheet differencesSheet = createDifferencesSheet(workbook, differencesSheetIndex, prodSheet);

            // Extract data from both sheets
            Set<String> sitData = extractSheetDataParallel(sitSheet);
            Set<String> prodData = extractSheetDataParallel(prodSheet);

            // Calculate differences between SIT and PROD, adding the source name at the end
            Set<String> diffData = sitData.stream()
                    .parallel()
                    .filter(row -> !prodData.contains(row))
                    .map(row -> row + "|SIT")  // Append the sheet name
                    .collect(Collectors.toSet());

            diffData.addAll(prodData.stream()
                    .parallel()
                    .filter(row -> !sitData.contains(row))
                    .map(row -> row + "|PROD")  // Append the sheet name
                    .collect(Collectors.toSet()));

            // Write differences to the sheet with source information
            writeDifferencesToSheet(differencesSheet, diffData);

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }

            System.out.println("Differences written to the 'Differences' sheet.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Sheet createDifferencesSheet(Workbook workbook, int index, Sheet prodSheet) {
        // Remove the "Differences" sheet if it already exists
        int sheetIndex = workbook.getSheetIndex("Differences");
        if (sheetIndex != -1) {
            workbook.removeSheetAt(sheetIndex);
        }

        // Create new "Differences" sheet at the specified index
        Sheet differencesSheet = workbook.createSheet("Differences");
        workbook.setSheetOrder("Differences", index);

        // Copy header from "PROD" sheet
        Row headerRow = prodSheet.getRow(0);
        Row newHeaderRow = differencesSheet.createRow(0);
        for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
            Cell oldCell = headerRow.getCell(i);
            Cell newCell = newHeaderRow.createCell(i);
            newCell.setCellValue(oldCell.toString());
        }

        // Add "Source Sheet" column header at the end
        newHeaderRow.createCell(headerRow.getPhysicalNumberOfCells()).setCellValue("Source Sheet");

        return differencesSheet;
    }

    private static Set<String> extractSheetDataParallel(Sheet sheet) {
        int rowCount = sheet.getPhysicalNumberOfRows();
        int columnCount = sheet.getRow(0).getPhysicalNumberOfCells();

        ConcurrentMap<Integer, String> dataMap = new ConcurrentHashMap<>();

        IntStream.range(1, rowCount).parallel().forEach(i -> {
            StringBuilder rowData = new StringBuilder();
            Row row = sheet.getRow(i);
            for (int j = 0; j < columnCount; j++) {
                Cell cell = row.getCell(j);
                rowData.append(cell.toString()).append("|");
            }
            dataMap.put(i, rowData.toString());
        });

        return dataMap.values().stream().collect(Collectors.toSet());
    }

    private static void writeDifferencesToSheet(Sheet sheet, Set<String> diffData) {
        int rowIndex = 1; // Start from the second row, since the first row is the header
        for (String row : diffData) {
            Row newRow = sheet.createRow(rowIndex++);
            String[] cells = row.split("\\|");

            // Write data cells
            for (int i = 0; i < cells.length; i++) {
                newRow.createCell(i).setCellValue(cells[i]);
            }
        }
    }
}

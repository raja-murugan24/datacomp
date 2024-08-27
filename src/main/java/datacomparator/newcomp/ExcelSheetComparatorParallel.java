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
            Sheet differencesSheet = workbook.createSheet("Differences");

            // Explicitly declare Set<String> types
            Set<String> sitData = extractSheetDataParallel(sitSheet);
            Set<String> prodData = extractSheetDataParallel(prodSheet);

            // Calculate differences between SIT and PROD
            Set<String> diffData = sitData.stream()
                    .parallel()
                    .filter(row -> !prodData.contains(row))
                    .map(String.class::cast)  // Ensure the type is String
                    .collect(Collectors.toSet());

            // Calculate differences between PROD and SIT and add to diffData
            diffData.addAll(prodData.stream()
                    .parallel()
                    .filter(row -> !sitData.contains(row))
                    .map(String.class::cast)  // Ensure the type is String
                    .collect(Collectors.toSet()));

            writeDifferencesToSheet(differencesSheet, diffData);

            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }

            System.out.println("Differences written to the 'Differences' sheet.");

        } catch (IOException e) {
            e.printStackTrace();
        }
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

        // Ensure the returned set is of type Set<String>
        return dataMap.values().stream().collect(Collectors.toSet());
    }

    private static void writeDifferencesToSheet(Sheet sheet, Set<String> diffData) {
        int rowIndex = 0;
        for (String row : diffData) {
            Row newRow = sheet.createRow(rowIndex++);
            String[] cells = row.split("\\|");
            for (int i = 0; i < cells.length; i++) {
                newRow.createCell(i).setCellValue(cells[i]);
            }
        }
    }
}

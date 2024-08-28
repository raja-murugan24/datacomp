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
            writeDifferencesToSheet(differencesSheet, diffData, prodSheet);

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

        // Copy header from "PROD" sheet with formatting
        Row headerRow = prodSheet.getRow(0);
        Row newHeaderRow = differencesSheet.createRow(0);
        copyRowWithFormatting(headerRow, newHeaderRow);

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

    private static void writeDifferencesToSheet(Sheet sheet, Set<String> diffData, Sheet prodSheet) {
        int rowIndex = 1; // Start from the second row, since the first row is the header
        for (String row : diffData) {
            Row newRow = sheet.createRow(rowIndex++);
            String[] cells = row.split("\\|");

            // Copy data cells with formatting from PROD sheet
            for (int i = 0; i < cells.length - 1; i++) { // Exclude last element, which is the sheet name
                Cell prodCell = prodSheet.getRow(1).getCell(i); // Reference cell from PROD sheet (row 1)
                Cell newCell = newRow.createCell(i);

                // Set cell value with the corresponding type
                newCell.setCellValue(cells[i]);
                if (prodCell != null) {
                    copyCellStyle(prodCell, newCell);
                }
            }

            // Set "Source Sheet" at the last column
            newRow.createCell(cells.length - 1).setCellValue(cells[cells.length - 1]);
        }
    }

    private static void copyRowWithFormatting(Row sourceRow, Row targetRow) {
        for (int i = 0; i < sourceRow.getPhysicalNumberOfCells(); i++) {
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell = targetRow.createCell(i);
            if (oldCell != null) {
                newCell.setCellValue(oldCell.toString());
                copyCellStyle(oldCell, newCell);
            }
        }
    }

    private static void copyCellStyle(Cell sourceCell, Cell targetCell) {
        Workbook workbook = sourceCell.getSheet().getWorkbook();
        CellStyle newCellStyle = workbook.createCellStyle();
        newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
        targetCell.setCellStyle(newCellStyle);
    }
}

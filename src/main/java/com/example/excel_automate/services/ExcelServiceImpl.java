package com.example.excel_automate.services;

import com.example.excel_automate.WorkbookClient;
import com.example.excel_automate.models.LanguageType;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelServiceImpl{

    public LanguageType languageType = new LanguageType();

//    @Override
    public Workbook deleteUnwantedColumns(Workbook workbook, String lang) {
        // Get the required columns based on the language
        List<String> requiredLangColumns = languageType.addLanguagesBasedOnCondition(lang);

        System.out.println("Required Columns: " + requiredLangColumns);  // Print out required columns for debugging

        // Iterate through each sheet in the workbook
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            System.out.println("Processing sheet: " + sheet.getSheetName());

            // Get the header row (row 0)
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                continue; // Skip empty sheets
            }

            // Collect column indices to delete
            List<Integer> columnsToDelete = new ArrayList<>();
            for (int cellIndex = 0; cellIndex < headerRow.getLastCellNum(); cellIndex++) {
                String headerValue = getCellValue(headerRow.getCell(cellIndex));

                // Print out the column header values for debugging
                System.out.println("Checking header value: " + headerValue);

                // If the column is not required, add it to the delete list
                if (!requiredLangColumns.contains(headerValue)) {
                    columnsToDelete.add(cellIndex);
                }
            }

            // Print out the columns that will be deleted
            System.out.println("Columns to delete: " + columnsToDelete);

            // Delete columns in reverse order to avoid shifting issues
            for (int i = columnsToDelete.size() - 1; i >= 0; i--) {
                int columnIndex = columnsToDelete.get(i);
                deleteColumn(sheet, columnIndex);
            }

            // Now shift all the remaining cells to the left to remove gaps
            shiftColumnsLeft(sheet);
        }

        return workbook; // Return the updated workbook
    }

    // Shifting the remaining cells to the left after column deletions
    private void shiftColumnsLeft(Sheet sheet) {
        // Iterate over all rows
        for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                int lastCellIndex = row.getPhysicalNumberOfCells();
                int cellIndex = 0;

                // Shift cells to the left if there are any gaps (empty cells)
                while (cellIndex < lastCellIndex) {
                    Cell currentCell = row.getCell(cellIndex);

                    if (currentCell == null || currentCell.toString().trim().isEmpty()) {
                        // Find the next non-empty cell
                        int nextCellIndex = findNextNonEmptyCell(row, cellIndex + 1);

                        if (nextCellIndex != -1) {
                            // Shift the next cell value to the current cell position
                            Cell nextCell = row.getCell(nextCellIndex);
                            row.createCell(cellIndex).setCellValue(nextCell.toString());

                            // Remove the shifted cell
                            row.removeCell(nextCell);
                        }
                    }
                    cellIndex++;
                }
            }
        }
    }

    // Find the next non-empty cell in the row starting from the given index
    private int findNextNonEmptyCell(Row row, int startIndex) {
        for (int i = startIndex; i < row.getPhysicalNumberOfCells(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && !cell.toString().trim().isEmpty()) {
                return i;
            }
        }
        return -1; // No non-empty cell found
    }

    private void deleteColumn(Sheet sheet, int colIndex) {
        // Iterate through all rows and shift left all cells to the right of the deleted column
        int totalRows = sheet.getPhysicalNumberOfRows();
        for (int rowIndex = 0; rowIndex < totalRows; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                // Shift cells to the left after deleting the column
                shiftCellsLeft(row, colIndex);
            }
        }
    }

    private void shiftCellsLeft(Row row, int colIndex) {
        // Get the total number of cells in the row
        int lastCellNum = row.getPhysicalNumberOfCells();

        // Shift all cells from colIndex + 1 to the left
        for (int i = colIndex + 1; i < lastCellNum; i++) {
            Cell currentCell = row.getCell(i);
            if (currentCell != null) {
                // Shift the cell value to the left cell
                Cell leftCell = row.createCell(i - 1);
                leftCell.setCellValue(currentCell.toString());

                // Clear the original cell
                row.removeCell(currentCell);
            }
        }

        // After shifting, remove the last cell in the row
        Cell lastCell = row.getCell(lastCellNum - 1);
        if (lastCell != null) {
            row.removeCell(lastCell);
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    // New method to process multiple languages
    // Process the workbook for multiple languages
    public void processMultipleLanguages(String inputFilePath, List<String> languages) {
        for (String language : languages) {
            System.out.println("Processing language: " + language);

            // Reload the workbook from the original file for each language
            try (FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
                Workbook workbook = WorkbookFactory.create(fis);

                // Process the workbook for the current language
                Workbook modifiedWorkbook = deleteUnwantedColumns(workbook, language);

                // Save the modified workbook for the current language
                String outputFilePath = "C:\\PowerAutomate\\modified_" + language.replaceAll("[^a-zA-Z0-9]", "_") + ".xlsm";
                try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                    modifiedWorkbook.write(fos);
                    modifiedWorkbook.close();
                    System.out.println("File saved for language: " + language);
                }
            } catch (IOException e) {
                e.printStackTrace();
                System.err.println("An error occurred while processing the language: " + language);
            }
        }
    }

}

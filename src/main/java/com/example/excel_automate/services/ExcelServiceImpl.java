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
        for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                int lastCellNum = row.getLastCellNum();
                if (colIndex >= 0 && colIndex < lastCellNum) {
                    for (int i = colIndex; i < lastCellNum - 1; i++) {
                        Cell oldCell = row.getCell(i + 1);
                        Cell newCell = row.getCell(i);
                        if (newCell == null) {
                            newCell = row.createCell(i);
                        }
                        if (oldCell != null) {
                            newCell.setCellValue(getCellValue(oldCell));
                        } else {
                            newCell.setBlank();
                        }
                    }
                    // Remove the last cell after shifting
                    Cell lastCell = row.getCell(lastCellNum - 1);
                    if (lastCell != null) {
                        row.removeCell(lastCell);
                    }
                }
            }
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
                copyOfDifferentDisplay(modifiedWorkbook,256,277,7);

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

    public void copyOfDifferentDisplay(Workbook workbook, int startRow, int endRow, int sourceSheetId) {
        // Get the source sheet using the provided sheet ID
        Sheet sourceSheet = workbook.getSheetAt(sourceSheetId);

        // Iterate through the target sheets (2 to 8, corresponding to indices 1 to 7)
        for (int sheetIndex = 1; sheetIndex <= 6; sheetIndex++) {
            Sheet targetSheet = workbook.getSheetAt(sheetIndex);

            // Copy rows from the specified range
            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                Row sourceRow = sourceSheet.getRow(rowIndex);
                if (sourceRow != null) {
                    // Create the corresponding row in the target sheet
                    Row targetRow = targetSheet.createRow(rowIndex);

                    // Copy each cell from the source row to the target row
                    for (int colIndex = 0; colIndex < sourceRow.getLastCellNum(); colIndex++) {
                        Cell sourceCell = sourceRow.getCell(colIndex);
                        if (sourceCell != null) {
                            Cell targetCell = targetRow.createCell(colIndex);

                            // Copy value and type from sourceCell to targetCell
                            switch (sourceCell.getCellType()) {
                                case STRING:
                                    targetCell.setCellValue(sourceCell.getStringCellValue());
                                    break;
                                case NUMERIC:
                                    targetCell.setCellValue(sourceCell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    targetCell.setCellValue(sourceCell.getBooleanCellValue());
                                    break;
                                case FORMULA:
                                    targetCell.setCellFormula(sourceCell.getCellFormula());
                                    break;
                                default:
                                    targetCell.setCellValue(sourceCell.toString());
                                    break;
                            }

                            // Optionally copy the cell style
                            targetCell.setCellStyle(sourceCell.getCellStyle());
                        }
                    }
                }
            }
        }

    }
}

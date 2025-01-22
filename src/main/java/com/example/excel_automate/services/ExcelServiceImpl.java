package com.example.excel_automate.services;

import com.example.excel_automate.dtos.RequestDto;
import com.example.excel_automate.models.LanguageType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.List;

public class ExcelServiceImpl implements ExcelService {

    public LanguageType languageType = new LanguageType();

    @Override
    public Workbook deleteUnwatedColumns(Workbook workbook, String lang) {
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

            // Now remove empty cells by shifting remaining cells left
            removeEmptyCellsAndShift(sheet);
        }

        return workbook; // Return the updated workbook
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

    private void removeEmptyCellsAndShift(Sheet sheet) {
        int totalRows = sheet.getPhysicalNumberOfRows();  // Get the number of rows

        // Iterate through each row
        for (int rowIndex = 0; rowIndex < totalRows; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                int lastCellIndex = row.getPhysicalNumberOfCells();  // Get the last cell index in the row

                // Iterate backward to avoid skipping cells after shifting
                for (int cellIndex = 0; cellIndex < lastCellIndex; cellIndex++) {
                    Cell currentCell = row.getCell(cellIndex);

                    // Check if the current cell is null or empty
                    if (currentCell == null || currentCell.toString().trim().isEmpty()) {
                        // Check if there are any non-empty cells from the current index to the end of the row
                        boolean hasNonEmptyCellAfter = false;

                        // Look ahead in the row from the current cell position until the last cell
                        for (int checkIndex = cellIndex + 1; checkIndex < lastCellIndex; checkIndex++) {
                            Cell checkCell = row.getCell(checkIndex);
                            if (checkCell != null && !checkCell.toString().trim().isEmpty()) {
                                hasNonEmptyCellAfter = true;
                                break;  // We found a non-empty cell after, no need to shift
                            }
                        }

                        // Only shift if there is a non-empty cell after the current one
                        if (hasNonEmptyCellAfter) {
                            for (int i = cellIndex; i < lastCellIndex - 1; i++) {
                                Cell nextCell = row.getCell(i + 1);

                                // Only shift if the next cell exists
                                if (nextCell != null) {
                                    Cell cellToShift = row.createCell(i);
                                    cellToShift.setCellValue(nextCell.toString());
                                }
                            }

                            // Remove the last cell in the row after shifting
                            Cell lastCell = row.getCell(lastCellIndex - 1);
                            if (lastCell != null) {
                                row.removeCell(lastCell);  // Remove the last cell as it is now redundant
                            }

                            break;  // Skip to the next row after handling the empty cell
                        }
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
}

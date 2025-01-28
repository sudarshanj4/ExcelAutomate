package com.example.excel_automate.services;

import com.example.excel_automate.models.FolderNaming;
import com.example.excel_automate.models.LanguageType;
import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class ExcelServiceImpl {

    public LanguageType languageType = new LanguageType();
    public FolderNaming folderNaming = new FolderNaming();

    public Workbook deleteUnwantedColumns(Workbook workbook, String lang) {
        List<String> requiredLangColumns = languageType.addLanguagesBasedOnCondition(lang);

        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            String sheetName = sheet.getSheetName();

            boolean isSpecialSheet = sheetName.equals("Delta_to_X2Pro10") || sheetName.equals("Delta_to_X2Extreme12");

            Row headerRow = sheet.getRow(0);
            if (headerRow == null) continue;

            List<Integer> columnsToDelete = new ArrayList<>();
            for (int cellIndex = 0; cellIndex < headerRow.getLastCellNum(); cellIndex++) {
                String headerValue = getCellValue(headerRow.getCell(cellIndex));
                if (!requiredLangColumns.contains(headerValue)) {
                    columnsToDelete.add(cellIndex);
                }
            }

            if (isSpecialSheet) {
                handleSpecialSheetDeletion(sheet, columnsToDelete);
            } else {
                deleteColumns(sheet, columnsToDelete);
                shiftColumnsLeft(sheet);
            }

            // Copy text from Column F to columns B to E at row 2500 for non-standard languages
            if (!lang.equalsIgnoreCase("Standard")) {
                copyTextFromColumnFToSpecificColumns(sheet);
            }
        }

        return workbook;
    }

    private void copyTextFromColumnFToSpecificColumns(Sheet sheet) {
        Row sourceRow = sheet.getRow(2500);
        if (sourceRow == null) return;

        Cell sourceCell = sourceRow.getCell(6); // Column G is index 6
        if (sourceCell == null || sourceCell.toString().trim().isEmpty()) return;

        String textToCopy = getCellValue(sourceCell);

        // Copy the text to columns B to E (indexes 1 to 5) at row 2500
        for (int colIndex = 1; colIndex <= 5; colIndex++) {
            Cell targetCell = sourceRow.getCell(colIndex);
            if (targetCell == null) {
                targetCell = sourceRow.createCell(colIndex);
            }
            targetCell.setCellValue(textToCopy);
        }
    }

    private void handleSpecialSheetDeletion(Sheet sheet, List<Integer> columnsToDelete) {
        for (int colIndex : columnsToDelete) {
            for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    Cell cell = row.getCell(colIndex);
                    if (cell != null) {
                        row.removeCell(cell);
                    }
                }
            }
        }

        shiftColumnsAndRemoveGaps(sheet);
    }

    private void shiftColumnsAndRemoveGaps(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        int lastColNum = sheet.getRow(0).getLastCellNum();

        for (int rowIndex = 0; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                for (int colIndex = 0; colIndex < lastColNum; colIndex++) {
                    Cell currentCell = row.getCell(colIndex);
                    if (currentCell == null || currentCell.toString().trim().isEmpty()) {
                        int nextNonEmptyIndex = findNextNonEmptyCell(row, colIndex + 1);
                        if (nextNonEmptyIndex != -1) {
                            Cell nextCell = row.getCell(nextNonEmptyIndex);
                            if (nextCell != null) {
                                row.createCell(colIndex).setCellValue(getCellValue(nextCell));
                                row.removeCell(nextCell);
                            }
                        }
                    }
                }
            }
        }
    }

    private int findNextNonEmptyCell(Row row, int startIndex) {
        for (int i = startIndex; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && !cell.toString().trim().isEmpty()) {
                return i;
            }
        }
        return -1;
    }

    private void deleteColumns(Sheet sheet, List<Integer> columnsToDelete) {
        for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                for (int i = columnsToDelete.size() - 1; i >= 0; i--) {
                    int colIndex = columnsToDelete.get(i);
                    int lastCellNum = row.getLastCellNum();
                    if (colIndex >= 0 && colIndex < lastCellNum) {
                        for (int j = colIndex; j < lastCellNum - 1; j++) {
                            Cell currentCell = row.getCell(j);
                            Cell nextCell = row.getCell(j + 1);
                            if (currentCell == null) {
                                currentCell = row.createCell(j);
                            }
                            if (nextCell != null) {
                                currentCell.setCellValue(getCellValue(nextCell));
                                currentCell.setCellStyle(nextCell.getCellStyle());
                            } else {
                                currentCell.setBlank();
                            }
                        }
                        row.removeCell(row.getCell(lastCellNum - 1));
                    }
                }
            }
        }
    }

    private void shiftColumnsLeft(Sheet sheet) {
        for (int rowIndex = 0; rowIndex < sheet.getPhysicalNumberOfRows(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
                    Cell currentCell = row.getCell(colIndex);
                    if (currentCell == null || currentCell.toString().trim().isEmpty()) {
                        Cell nextCell = row.getCell(colIndex + 1);
                        if (nextCell != null) {
                            currentCell = row.createCell(colIndex);
                            currentCell.setCellValue(getCellValue(nextCell));
                            currentCell.setCellStyle(nextCell.getCellStyle());
                            row.removeCell(nextCell);
                        }
                    }
                }
            }
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                String stringValue = cell.getStringCellValue();
                // Check if the value already contains quotes
                if (stringValue.startsWith("\"") && stringValue.endsWith("\"")) {
                    return stringValue; // Keep the quotes if they exist
                } else {
                    return stringValue; // Return as is if no quotes
                }
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    public void processMultipleLanguages(String inputFilePath, List<String> languages, String version) {
        for (String language : languages) {
            try (FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
                Workbook workbook = WorkbookFactory.create(fis);
                deleteUnwantedColumns(workbook, language);

                // Process for source sheet ID 7 (add "10'")
                System.out.println("Processing source sheet ID 7...");
                copyOfDifferentDisplaySizes(workbook, 256, 277, 7);
                saveSheetsWithTextAddition(workbook, language, "10'",version);

                // Process for source sheet ID 8 (add "12'")
                System.out.println("Processing source sheet ID 8...");
                copyOfDifferentDisplaySizes(workbook, 256, 277, 8);
                saveSheetsWithTextAddition(workbook, language, "12'",version);

                // Create a folder for the selected language
                String languageFolderPath = "C:\\PowerAutomate\\Excel\\" + folderNaming.folderName(language);
                createFolder(languageFolderPath);

                // Save the modified file in the language-specific folder
                String outputFilePath = languageFolderPath + "\\modified_" + language.replaceAll("[^a-zA-Z0-9]", "_") + ".xlsm";
                try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                    workbook.write(fos);
                }

                workbook.close();
            } catch (IOException e) {
                System.err.println("Error processing language " + language + ": " + e.getMessage());
                e.printStackTrace();
            }
        }
    }

    private String cellToString(Cell cell) {
        if (cell == null) {
            return ""; // Return empty for null cells
        }

        // Use DataFormatter to get the cell's displayed value as-is
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    public void saveSheetsWithTextAddition(Workbook workbook, String language, String additionalText, String versionno) {
        // Determine the displaySize (_10 or _12) based on additionalText
        String displaySize = "";
        if ("10'".equals(additionalText)) {
            displaySize = "X2Pro10";
        } else if ("12'".equals(additionalText)) {
            displaySize = "X2Extreme12";
        }

        for (int sheetIndex = 1; sheetIndex < workbook.getNumberOfSheets() - 3; sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            String sheetName = sheet.getSheetName();
            String languageFolderPath = "C:\\PowerAutomate\\Excel\\" + folderNaming.folderName(language);

            // Ensure the folder exists
            File languageFolder = new File(languageFolderPath);
            if (!languageFolder.exists()) {
                boolean dirCreated = languageFolder.mkdirs();
                if (!dirCreated) {
                    System.err.println("Failed to create folder: " + languageFolderPath);
                    continue; // Skip to the next sheet if folder creation fails
                }
            }

            // Create the text file for the current sheet with displaySize (_10 or _12)
            File textFile = new File(languageFolderPath + "\\" + "_Txt" + folderNaming.folderName(language) + "_" + displaySize + "_S1_" + versionno + "_" + sheetName + ".txt");
            try (BufferedWriter writer = new BufferedWriter(new FileWriter(textFile, true))) { // 'true' enables append mode
                for (Row row : sheet) {
                    List<String> cellValues = new ArrayList<>();
                    for (Cell cell : row) {
                        // Simply get the cell value as a string
                        cellValues.add(cellToString(cell));
                    }
                    // Write the row as tab-separated values without adding extra quotes
                    writer.write(String.join("\t", cellValues));
                    writer.newLine(); // Add a new line for each row
                }



                System.out.println("Saved text file: " + textFile.getAbsolutePath());
            } catch (IOException e) {
                System.err.println("Error writing sheet " + sheetName + " to text file: " + e.getMessage());
                e.printStackTrace();
            }
        }
    }

    public void copyOfDifferentDisplaySizes(Workbook workbook, int startRow, int endRow, int sourceSheetId) {
        Sheet sourceSheet = workbook.getSheetAt(sourceSheetId);

        for (int sheetIndex = 1; sheetIndex <= 6; sheetIndex++) {
            Sheet targetSheet = workbook.getSheetAt(sheetIndex);
            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
                Row sourceRow = sourceSheet.getRow(rowIndex);
                if (sourceRow != null) {
                    Row targetRow = targetSheet.createRow(rowIndex);
                    for (int colIndex = 0; colIndex < sourceRow.getLastCellNum(); colIndex++) {
                        Cell sourceCell = sourceRow.getCell(colIndex);
                        if (sourceCell != null) {
                            Cell targetCell = targetRow.createCell(colIndex);
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
                            targetCell.setCellStyle(sourceCell.getCellStyle());
                        }
                    }
                }
            }
        }
    }

    private void createFolder(String folderPath) {
        File folder = new File(folderPath);
        if (!folder.exists()) {
            folder.mkdirs();
        }
    }
}

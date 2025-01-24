package com.example.excel_automate.services;

import com.example.excel_automate.models.LanguageType;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelServiceImpl {

    public LanguageType languageType = new LanguageType();

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
        }

        return workbook;
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
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf((int) cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    public void processMultipleLanguages(String inputFilePath, List<String> languages) {
        for (String language : languages) {
            try (FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
                Workbook workbook = WorkbookFactory.create(fis);
                deleteUnwantedColumns(workbook, language);
                copyOfDifferentDisplay(workbook, 256, 277, 7);

                String outputFilePath = "C:\\PowerAutomate\\modified_" + language.replaceAll("[^a-zA-Z0-9]", "_") + ".xlsm";
                try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                    workbook.write(fos);
                    workbook.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void copyOfDifferentDisplay(Workbook workbook, int startRow, int endRow, int sourceSheetId) {
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
}

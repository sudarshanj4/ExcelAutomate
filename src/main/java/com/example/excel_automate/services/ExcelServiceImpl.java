package com.example.excel_automate.services;

import com.example.excel_automate.models.LanguageType;
import com.example.excel_automate.services.ExcelService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelServiceImpl implements ExcelService {

    public LanguageType languageType = new LanguageType();

//    @Override
    public void processMultipleLanguages(Workbook mainWorkbook, List<String> languages) throws IOException {
        for (String language : languages) {
            System.out.println("Processing for language: " + language);

            Workbook workbookCopy = createWorkbookCopy(mainWorkbook);

            deleteUnwatedColumns(workbookCopy, language);

            saveWorkbook(workbookCopy, language);

            workbookCopy.close();
        }
    }

    private Workbook createWorkbookCopy(Workbook sourceWorkbook) {
        Workbook newWorkbook = new XSSFWorkbook();
        for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {
            Sheet sourceSheet = sourceWorkbook.getSheetAt(i);
            Sheet targetSheet = newWorkbook.createSheet(sourceSheet.getSheetName());
            copySheetContents(sourceSheet, targetSheet, sourceWorkbook, newWorkbook);
        }
        return newWorkbook;
    }

    private void copySheetContents(Sheet sourceSheet, Sheet targetSheet, Workbook sourceWorkbook, Workbook targetWorkbook) {
        Map<CellStyle, CellStyle> styleCache = new HashMap<>();
        for (int rowIndex = 0; rowIndex <= sourceSheet.getLastRowNum(); rowIndex++) {
            Row sourceRow = sourceSheet.getRow(rowIndex);
            if (sourceRow == null) continue;

            Row targetRow = targetSheet.createRow(rowIndex);
            targetRow.setHeight(sourceRow.getHeight());

            for (int cellIndex = 0; cellIndex < sourceRow.getLastCellNum(); cellIndex++) {
                Cell sourceCell = sourceRow.getCell(cellIndex);
                if (sourceCell == null) continue;

                Cell targetCell = targetRow.createCell(cellIndex);
                copyCell(sourceCell, targetCell, targetWorkbook, styleCache);
            }
        }
    }

    private void copyCell(Cell sourceCell, Cell targetCell, Workbook targetWorkbook, Map<CellStyle, CellStyle> styleCache) {
        targetCell.setCellType(sourceCell.getCellType());
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
                break;
        }
        if (sourceCell.getCellStyle() != null) {
            CellStyle cachedStyle = styleCache.get(sourceCell.getCellStyle());
            if (cachedStyle == null) {
                cachedStyle = targetWorkbook.createCellStyle();
                cachedStyle.cloneStyleFrom(sourceCell.getCellStyle());
                styleCache.put(sourceCell.getCellStyle(), cachedStyle);
            }
            targetCell.setCellStyle(cachedStyle);
        }
    }

    public Workbook deleteUnwatedColumns(Workbook workbook, String lang) {
        List<String> requiredLangColumns = languageType.addLanguagesBasedOnCondition(lang);
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) continue;

            List<Integer> columnsToDelete = new ArrayList<>();
            for (int cellIndex = 0; cellIndex < headerRow.getLastCellNum(); cellIndex++) {
                String headerValue = getCellValue(headerRow.getCell(cellIndex));
                if (!requiredLangColumns.contains(headerValue)) {
                    columnsToDelete.add(cellIndex);
                }
            }
            for (int i = columnsToDelete.size() - 1; i >= 0; i--) {
                deleteColumn(sheet, columnsToDelete.get(i));
            }
        }
        return workbook;
    }

    private void deleteColumn(Sheet sheet, int colIndex) {
        for (Row row : sheet) {
            if (row != null) {
                Cell cell = row.getCell(colIndex);
                if (cell != null) {
                    row.removeCell(cell);
                }
            }
        }
    }

    private void saveWorkbook(Workbook workbook, String language) throws IOException {
        String destinationFolder = "C:\\PowerAutomate";
        File folder = new File(destinationFolder);
        if (!folder.exists() && !folder.mkdirs()) {
            throw new IOException("Failed to create destination folder: " + destinationFolder);
        }
        String fileName = destinationFolder + "\\output_" + language.replaceAll("[^a-zA-Z0-9]", "_") + ".xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
            System.out.println("Workbook saved as: " + fileName);
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

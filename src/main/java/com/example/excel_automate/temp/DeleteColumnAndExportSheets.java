package com.example.excel_automate.temp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class DeleteColumnAndExportSheets {

    public static void main(String[] args) {
        String filePath = "C:\\PowerAutomate\\X2_Collect_V3_01_16_02_00_.xlsm";

        // Record the start time
        long startTime = System.nanoTime();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Get the directory of the input file
            File inputFile = new File(filePath);
            String outputDirectory = inputFile.getParent();

            // Iterate through all sheets in the workbook
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);

                // Remove column K (index 10, zero-based) from the sheet
                deleteColumn(sheet, 10);

                // Export the sheet to a .txt file in the same directory as the input file
                String sheetName = sheet.getSheetName().replaceAll("[\\\\/:*?\"<>|]", "_"); // Sanitize the sheet name
                String outputFilePath = outputDirectory + File.separator + sheetName + ".txt";
                exportSheetToTxt(sheet, outputFilePath);
            }

            // Save the updated workbook
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
                System.out.println("Column K deleted successfully from all sheets.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        // Record the end time
        long endTime = System.nanoTime();

        // Calculate the duration in milliseconds
        long duration = (endTime - startTime) / 1_000_000; // Convert to milliseconds
        System.out.println("Time taken to complete the task: " + duration + " ms");
    }

    private static void deleteColumn(Sheet sheet, int colIndex) {
        for (Row row : sheet) {
            if (row != null && row.getLastCellNum() > colIndex) {
                // Shift cells to the left starting from colIndex
                for (int i = colIndex; i < row.getLastCellNum() - 1; i++) {
                    Cell currentCell = row.getCell(i);
                    Cell nextCell = row.getCell(i + 1);

                    if (currentCell == null) {
                        currentCell = row.createCell(i);
                    }

                    if (nextCell != null) {
                        // Handle different cell types
                        switch (nextCell.getCellType()) {
                            case STRING -> currentCell.setCellValue(nextCell.getStringCellValue());
                            case NUMERIC -> currentCell.setCellValue(nextCell.getNumericCellValue());
                            case BOOLEAN -> currentCell.setCellValue(nextCell.getBooleanCellValue());
                            case FORMULA -> currentCell.setCellFormula(nextCell.getCellFormula());
                            case BLANK -> currentCell.setBlank();
                            default -> currentCell.setBlank(); // Handle other cases
                        }
                        currentCell.setCellStyle(nextCell.getCellStyle());
                    } else {
                        // Clear the current cell if there's no next cell
                        row.removeCell(currentCell);
                    }
                }
                // Remove the last cell explicitly
                Cell lastCell = row.getCell(row.getLastCellNum() - 1);
                if (lastCell != null) {
                    row.removeCell(lastCell);
                }
            }
        }
    }

    private static void exportSheetToTxt(Sheet sheet, String fileName) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(fileName))) {
            for (Row row : sheet) {
                StringBuilder rowContent = new StringBuilder();

                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING -> rowContent.append(cell.getStringCellValue()).append("\t");
                        case NUMERIC -> rowContent.append(cell.getNumericCellValue()).append("\t");
                        case BOOLEAN -> rowContent.append(cell.getBooleanCellValue()).append("\t");
                        case FORMULA -> rowContent.append(cell.getCellFormula()).append("\t");
                        case BLANK -> rowContent.append("\t");
                        default -> rowContent.append("\t");
                    }
                }

                // Write the row content to the file
                writer.write(rowContent.toString().trim());
                writer.newLine();
            }
            System.out.println("Exported sheet '" + sheet.getSheetName() + "' to " + fileName);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
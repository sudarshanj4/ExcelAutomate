package com.example.excel_automate.models;

import java.util.List;

public class ExcelWorkbook {
    private String workbookName;
    private String filePath;
    private List<String> sheetNames;

    // Constructor to initialize workbookName and filePath
    public ExcelWorkbook(String workbookName, String filePath) {
        this.workbookName = workbookName;
        this.filePath = filePath;
    }

    // Getter methods
    public String getWorkbookName() {
        return workbookName;
    }

    public String getFilePath() {
        return filePath;
    }

    public List<String> getSheetNames() {
        return sheetNames;
    }

    // Set the sheet names after processing the workbook
    public void setSheetNames(List<String> sheetNames) {
        this.sheetNames = sheetNames;
    }

    // Method to load sheet names from the workbook (no print statements here)
    public void loadSheetNames() {
        // Assuming WorkbookHandler is used to extract sheet names and populate
        WorkbookHandler workbookHandler = new WorkbookHandler();
        this.sheetNames = workbookHandler.getSheetNamesFromWorkbook(filePath);  // Example call to get sheet names
    }
}

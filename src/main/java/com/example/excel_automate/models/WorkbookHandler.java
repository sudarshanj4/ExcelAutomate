package com.example.excel_automate.models;

import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class WorkbookHandler {

    // Method to extract sheet names from the provided workbook file path
    public List<String> getSheetNamesFromWorkbook(String filePath) {
        List<String> sheetNames = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {
            // Loop through sheets and collect sheet names
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                sheetNames.add(sheet.getSheetName());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return sheetNames;
    }
}

package com.example.excel_automate;

import com.example.excel_automate.services.ExcelServiceImpl;

import java.util.List;

public class WorkbookClient {

    public static void main(String[] args) {
        // Define the input file path
        String inputFilePath = "C:\\PowerAutomate\\X2_Collect_V3_01_16_02_00_.xlsm"; // Input macro-enabled workbook

        // Define a list of languages for testing
        List<String> languages = List.of("Chinese (Simplified, China)", "Italian (Italy)"); // Multiple languages

        // Create an instance of the ExcelServiceImpl class
        ExcelServiceImpl excelService = new ExcelServiceImpl();

        System.out.println("Processing workbook...");

        // Process the workbook for each language
        excelService.processMultipleLanguages(inputFilePath, languages);

        System.out.println("Processing completed.");
    }
}

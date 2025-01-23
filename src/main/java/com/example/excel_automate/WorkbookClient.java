package com.example.excel_automate;

import com.example.excel_automate.services.ExcelServiceImpl;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

public class WorkbookClient {

    public static void main(String[] args) {
        // Define the input file path
        String inputFilePath = "C:\\PowerAutomate\\X2_Collect_V3_01_16_02_00_.xlsm"; // Input macro-enabled workbook

        // Define a list of languages for testing
        List<String> languages = List.of("Chinese (Simplified, China)", "Italian (Italy)"); // Multiple languages

        // Create an instance of the ExcelServiceImpl class
        ExcelServiceImpl excelService = new ExcelServiceImpl();

        // Open the workbook, process it for each language, and save each as a new file
        try (FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
            // Load the main workbook
            Workbook workbook = WorkbookFactory.create(fis);

            System.out.println("Processing workbook...");

            // Process the workbook for each language
            excelService.processMultipleLanguages(workbook, languages);

        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("An error occurred while processing the Excel file.");
        }
    }
}

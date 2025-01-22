package com.example.excel_automate;


import com.example.excel_automate.models.ExcelWorkbook;
import com.example.excel_automate.services.ExcelServiceImpl;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WorkbookClient {

    public static void main(String[] args) {
        // Define the input and output file paths
        String inputFilePath = "C:\\PowerAutomate\\X2_Collect_V3_01_16_02_00_.xlsm"; // Input macro-enabled workbook
        String outputFilePath = "C:\\PowerAutomate\\file.xlsm"; // Output macro-enabled workbook

        // Define the language for testing
        String language = "Chinese (Simplified, China)"; // Replace with the desired language

        // Create an instance of the ExcelServiceImpl class
        ExcelServiceImpl excelService = new ExcelServiceImpl();

        // Open the workbook, process it, and save the updated file
        try (FileInputStream fis = new FileInputStream(new File(inputFilePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            System.out.println("Processing workbook...");

            // Call the deleteUnwatedColumns method
            Workbook updatedWorkbook = excelService.deleteUnwatedColumns(workbook, language);

            // Save the updated workbook as .xlsm
            try (FileOutputStream fos = new FileOutputStream(new File(outputFilePath))) {
                updatedWorkbook.write(fos);
                System.out.println("Workbook updated and saved to: " + outputFilePath);
            }

        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("An error occurred while processing the Excel file.");
        }
    }
}



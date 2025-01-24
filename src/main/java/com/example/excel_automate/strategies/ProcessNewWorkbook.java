//package com.example.excel_automate.strategies;
//
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.ss.usermodel.WorkbookFactory;
//
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.IOException;
//
//public class ProcessNewWorkbook {
//    // Open the workbook, process it for each language, and save each as a new file
//    public void process(String inputFilePath) {
//        try (
//                FileInputStream fis = new FileInputStream(new File(inputFilePath))) {
//            // Load the main workbook
//            Workbook workbook = WorkbookFactory.create(fis);
//
//            System.out.println("Processing workbook...");
//
//            // Process the workbook for each language
//            excelService.processMultipleLanguages(workbook, languages);
//
//        } catch (
//                IOException e) {
//            e.printStackTrace();
//            System.err.println("An error occurred while processing the Excel file.");
//        }
//    }
//
//}

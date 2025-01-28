//package com.example.excel_automate;
//
//import com.example.excel_automate.services.ExcelServiceImpl;
//
//import java.util.List;
//
//public class WorkbookClient {
//
//    public static void main(String[] args) {
//        // Define the input file path
//        String inputFilePath = "C:\\PowerAutomate\\X2_Collect_V3_01_17_01_00_.xlsm"; // Input macro-enabled workbook
//
//        // Define a list of languages for testing
//        List<String> languages = List.of(
//                "English (United States)",
//                "French (France)",
//                "Italian (Italy)",
//                "Russian (Russia)",
//                "Japanese (Japan)",
//                "Chinese (Simplified, China)",
//                "Chinese (Traditional, Taiwan)",
//                "Arabic (Oman)"
//        );
//
//        // Create an instance of the ExcelServiceImpl class
//        ExcelServiceImpl excelService = new ExcelServiceImpl();
//
//        // Start timing the execution
//        long startTime = System.nanoTime();
//
//        System.out.println("Processing workbook...");
//
//        String versionno= "V3_01_17_01";
//
//        // Process the workbook for each language
//        excelService.processMultipleLanguages(inputFilePath, languages,versionno);
//
//        // End timing the execution
//        long endTime = System.nanoTime();
//        double durationInSeconds = (endTime - startTime) / 1_000_000_000.0;
//
//        System.out.println("Processing completed.");
//        System.out.println("Total execution time: " + durationInSeconds + " seconds");
//
//
//        System.out.println(durationInSeconds + " seconds");
//    }
//}

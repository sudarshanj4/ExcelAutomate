//package com.example.excel_automate.controllers;
//
//import com.example.excel_automate.dtos.RequestDto;
//import com.example.excel_automate.services.ExcelService;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.web.bind.annotation.PostMapping;
//import org.springframework.web.bind.annotation.RequestMapping;
//import org.springframework.web.bind.annotation.RequestParam;
//import org.springframework.web.bind.annotation.RestController;
//
//import java.io.FileNotFoundException;
//
//@RestController
//@RequestMapping("/excel")
//public class ExcelController {
//    private ExcelService excelService;
//    private RequestDto requestDto;
//
//    @Autowired
////    public ExcelController(ExcelService excelService) {
////        this.excelService = excelService;
////    }
//
//
//    @PostMapping("/process")
//    public String processExcelFile(@RequestParam String filePath) throws FileNotFoundException {
//        // Call the service method to process the Excel file
////        return excelService.processExcelFile(filePath);
//    }
//    public RequestDto getRequestDto() {
//        return this.requestDto;
//    }
//}

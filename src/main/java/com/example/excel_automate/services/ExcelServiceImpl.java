package com.example.excel_automate.services;

import com.example.excel_automate.controllers.ExcelController;
import com.example.excel_automate.dtos.RequestDto;
import com.example.excel_automate.models.EngineeType;
import com.example.excel_automate.models.LanguageType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
public class ExcelServiceImpl implements ExcelService {
    public ExcelController excelController;
    public EngineeType engineeType;
    public LanguageType languageType;

    @Override
    public String processExcelFile(String filePath) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Iterator<Sheet> sheetItrator = workbook.sheetIterator();

            List<String> sheetname = new ArrayList<>();
            while (sheetItrator.hasNext()) {
                Sheet sheet = sheetItrator.next();
                sheetname.add(sheet.getSheetName());
            }

            RequestDto requestDto = excelController.getRequestDto();
            try {
                List<String> requiredLanguages = languageType.addLanguagesBasedOnCondition(requestDto.getLanguage_name());


            } catch (Exception e) {
                System.out.println("Error occured while processing excel file");
            }

//            Sheet sheet = workbook.getSheetAt(0);
//            Row row = sheet.getRow(0);
//            if (row == null) {
//                row = sheet.createRow(0);
//            }
//            Cell cell = row.createCell(0);
//            cell.setCellValue("Processed");
//
//            // Save the changes to the same file
            try (FileOutputStream fos = new FileOutputStream(new File(filePath))) {
                workbook.write(fos);
            }

            return filePath; // Return the file path after processing

        } catch (IOException e) {
            e.printStackTrace();
            return "Error processing the file.";

        }
    }
}

package com.example.excel_automate.controllers;

import com.example.excel_automate.dtos.RequestDto;
import com.example.excel_automate.dtos.ResponseDTO;
import com.example.excel_automate.services.ExcelServiceImpl;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.FileNotFoundException;

@RestController
@RequestMapping("/excel")
public class ExcelController {

    ExcelServiceImpl excelService=new ExcelServiceImpl();
    ResponseDTO responseDTO=new ResponseDTO();
    @PostMapping("/automate")
    public ResponseEntity<String> creationTxtFile(@RequestBody RequestDto requestDto) throws FileNotFoundException {
        excelService.processMultipleLanguages(requestDto.getFilePathUrl(),requestDto.getLanguage_name(),requestDto.getVersion(),requestDto.getDestination_filePathUrl());

        return ResponseEntity.ok("Successfully created file");
    }
}

package org.example.excelgenerator.controller;

import org.example.excelgenerator.dto.request.ExcelRequest;
import org.example.excelgenerator.service.ExcellGenerator;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;

@RestController
@RequestMapping("/download-excel")
public class ExcelGeneratorController {
    private final ExcellGenerator excelGeneratorService;

    @Autowired
    public ExcelGeneratorController(ExcellGenerator excelGeneratorService) {
        this.excelGeneratorService = excelGeneratorService;
    }

    @PostMapping("/generate")
    public ResponseEntity<byte[]> generateExcel(@RequestBody ExcelRequest request) {
        try {
            byte[] excelData = excelGeneratorService.generateExcel(request);

            HttpHeaders headers = new HttpHeaders();
            headers.add("Content-Disposition", "attachment; filename=data.xlsx");
            headers.add("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            return new ResponseEntity<>(excelData, headers, HttpStatus.OK);
        } catch (IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).build();
        }
    }
}

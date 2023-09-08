package com.example.controller;

import com.example.service.WriteExcelDemo;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
@RequestMapping("/api/v1")
public class FileController {

    @Autowired
    private WriteExcelDemo writeExcelDemo;

    @PostMapping("/download/filetxt")
    public ResponseEntity<Object> downloadFileTxt() throws IOException {

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment")
                .contentType(MediaType.TEXT_PLAIN)
                .body("Hello txt file");
    }

    @GetMapping("/download/writeExcel")
    public ResponseEntity<Object> writeExcel() {
        InputStreamResource file = new InputStreamResource(writeExcelDemo.createExcel());

        String fileName = "demo-excel.xlsx";

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName)
                .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .body(file);

    }
}

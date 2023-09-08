package com.example.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

@Service
public class WriteExcelDemo {

    public ByteArrayInputStream createExcel() {

        try {

        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");

        //Prepare data to be written as an Object[]
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"ID", "First Name", "Last Name", "Date of Birth"});
        data.put("2", new Object[] {1, "Amit", "Shukla", "30-12-2000"});
        data.put("3", new Object[] {2, "Lokesh", "Gupta", "10-12-1990"});
        data.put("4", new Object[] {3, "John", "Adwards", "15-1-1992"});
        data.put("5", new Object[] {4, "Brian", "Schultz", "20-5-2010"});

        Set<String> keySet = data.keySet();
        int  rowNum = 0;

        CellStyle style = createWarningColor(workbook);

        for (String key: keySet) {

            Row row = sheet.createRow(rowNum);
            Object[] objArr = data.get(key);
            int cellNum = 0;
            for (Object obj: objArr) {

                Cell cell = row.createCell(cellNum++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);

                    // style only for row 0
                    if (rowNum == 0) {
                        cell.setCellStyle(style);
                    }
                } else if (obj instanceof Integer){
                    cell.setCellValue((Integer) obj);

                    // style only for row 0
                    if (rowNum == 0) {
                        cell.setCellStyle(style);
                    }
                }
            }
            rowNum++;
        }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            out.close();

            return new ByteArrayInputStream(out.toByteArray());
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
        return null;
    }

//    public ByteArrayInputStream createExcel() {
//
//        try {
//        //Blank workbook
//        XSSFWorkbook workbook = new XSSFWorkbook();
//
//        //Create a blank sheet
//        XSSFSheet sheet = workbook.createSheet("Employee Data");
//
//        // row 0
//        Row row = sheet.createRow(0);
//
//        Cell cell = row.createCell(0);
//        cell.setCellValue("ID");
//        row.createCell(1).setCellValue("Name");
//        row.createCell(2).setCellValue("Date of Birth");
//
//        // row 1
//        Row row1 = sheet.createRow(1);
//
//        Cell cell1 = row1.createCell(0);
//        cell1.setCellValue(1);
//        row1.createCell(1).setCellValue("John Doe");
//        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd-MM-yyyy");
//        Date dateParse = simpleDateFormat.parse("30-12-2023");
//        String dateFormatted = simpleDateFormat.format(dateParse);
//        row1.createCell(2).setCellValue(dateFormatted);
//
//
////        try {
//            ByteArrayOutputStream out = new ByteArrayOutputStream();
//            workbook.write(out);
//            out.close();
//
//            return new ByteArrayInputStream(out.toByteArray());
//        } catch (Exception e) {
//            System.out.println(e.getMessage());
//        }
//        return null;
//    }

    public CellStyle createWarningColor(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        return style;
    }

}

package com.example.excel.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.FileOutputStream;

@SpringBootTest
class SimpleTest {

        @Test
        void test() {
        // Create a new Excel workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a sheet in the workbook
        Sheet sheet = workbook.createSheet("Sheet1");

        // Create a row and put some cells in it
        Row row = sheet.createRow(0);

        // Create a cell with a value
        Cell cell = row.createCell(0);
        cell.setCellValue("Hello");

        // Create another cell with a numeric value
        row.createCell(1).setCellValue(123);

        // Create a third cell with a formula
        row.createCell(2).setCellFormula("A2+B2");

        try {
            // Write the workbook to a file
            try (FileOutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
                workbook.write(fileOut);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // Close the workbook to release resources
            try {
                workbook.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        System.out.println("Excel file created successfully.");
    }

}


package com.pranav.email.sender.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Service
public class ExcelUpdateService {

    public synchronized void updateStatus(String excelPath, String sheetName,
                                          String emailColumn, String email, String status) {

        File excelFile = new File(excelPath);
        if (!excelFile.exists()) throw new RuntimeException("Excel file not found: " + excelPath);

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFile))) {

            Sheet sheet = workbook.getSheet(sheetName);
            Row headerRow = sheet.getRow(1);

            int emailCol = findCol(headerRow, emailColumn);
            int statusCol = findOrCreateCol(headerRow, "Mail Status");
            int timeCol = findOrCreateCol(headerRow, "Mail Sent Time");

            for (int i = 2; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell cell = row.getCell(emailCol);
                if (cell == null) continue;

                if (email.equalsIgnoreCase(cell.getStringCellValue().trim())) {
                    row.createCell(statusCol).setCellValue(status);
                    row.createCell(timeCol).setCellValue(
                        LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd-MM-yyyy hh.mm a"))
                    );
                    break;
                }
            }

            // Write to temp file
            File tempFile = new File(excelFile.getParent(), "temp_" + excelFile.getName());
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                workbook.write(fos);
            }

            // Copy temp to original safely
            Path source = tempFile.toPath();
            Path target = excelFile.toPath();
            Files.copy(source, target, StandardCopyOption.REPLACE_EXISTING);
            Files.deleteIfExists(tempFile.toPath());

        } catch (Exception e) {
            System.err.println("Could not update Excel: " + e.getMessage());
        }
    }

    private int findCol(Row row, String name) {
        for (Cell c : row) {
            if (c.getCellType() == CellType.STRING &&
                c.getStringCellValue().trim().equalsIgnoreCase(name.trim())) {
                return c.getColumnIndex();
            }
        }
        throw new RuntimeException("Column missing: " + name);
    }

    private int findOrCreateCol(Row row, String name) {
        for (Cell c : row) {
            if (c.getCellType() == CellType.STRING &&
                c.getStringCellValue().trim().equalsIgnoreCase(name.trim())) {
                return c.getColumnIndex();
            }
        }
        int idx = row.getLastCellNum();
        if (idx < 0) idx = 0;
        row.createCell(idx).setCellValue(name);
        return idx;
    }
}

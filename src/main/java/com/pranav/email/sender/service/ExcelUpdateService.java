package com.pranav.email.sender.service;

import com.pranav.email.sender.config.ExcelConstants;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@Service
public class ExcelUpdateService {

    public synchronized void updateStatus(String email, String status) {

        try (Workbook workbook =
                     new XSSFWorkbook(new FileInputStream(ExcelConstants.FILE_PATH))) {

            Sheet sheet = workbook.getSheet(ExcelConstants.SHEET_NAME);
            Row headerRow = sheet.getRow(1);

            int emailCol = findCol(headerRow, ExcelConstants.EMAIL_COLUMN);
            int statusCol = findCol(headerRow, ExcelConstants.STATUS);
            int timeCol = findCol(headerRow, ExcelConstants.MAIL_SENT_TIME);

            for (int i = 2; i <= sheet.getLastRowNum(); i++) {

                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell cell = row.getCell(emailCol);
                if (cell == null) continue;

                if (email.equalsIgnoreCase(cell.getStringCellValue().trim())) {

                    row.createCell(statusCol).setCellValue(status);
                    row.createCell(timeCol).setCellValue(
                            LocalDateTime.now().format(
                                    DateTimeFormatter.ofPattern("dd-MM-yyyy hh.mm a")));
                    break;
                }
            }

            try (FileOutputStream fos =
                         new FileOutputStream(ExcelConstants.FILE_PATH)) {
                workbook.write(fos);
            }

        } catch (Exception ignored) {}
    }

    private int findCol(Row row, String name) {
        for (Cell c : row) {
            if (c.getCellType() == CellType.STRING &&
                c.getStringCellValue().trim().equalsIgnoreCase(name)) {
                return c.getColumnIndex();
            }
        }
        throw new RuntimeException("Column missing");
    }
}

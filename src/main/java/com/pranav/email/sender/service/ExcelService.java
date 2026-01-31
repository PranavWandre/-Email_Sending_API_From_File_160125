package com.pranav.email.sender.service;

import com.pranav.email.sender.config.ExcelConstants;
import com.pranav.email.sender.dto.MailRequest;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

@Service
public class ExcelService {

    private final EmailService emailService;

    public ExcelService(EmailService emailService) {
        this.emailService = emailService;
    }

    public void sendMailsFromExcel(MailRequest request) {

        File excelFile = new File(request.getExcelFilePath());
        if (!excelFile.exists()) throw new RuntimeException("Excel file not found: " + request.getExcelFilePath());

        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFile))) {

            Sheet sheet = workbook.getSheet(request.getSheetName());
            Row headerRow = sheet.getRow(1);

            int emailCol = getColumnIndex(headerRow, request.getEmailColumn());
            int statusCol = getOrCreateColumn(headerRow, ExcelConstants.STATUS);
            int timeCol = getOrCreateColumn(headerRow, ExcelConstants.MAIL_SENT_TIME);

            for (int i = 2; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell emailCell = row.getCell(emailCol);
                if (emailCell == null || emailCell.getStringCellValue().trim().isEmpty()
                        || !emailCell.getStringCellValue().contains("@")) {
                    updateStatus(row, statusCol, timeCol, "INVALID", false);
                    continue;
                }

                String email = emailCell.getStringCellValue().trim();

                emailService.sendMail(
                        email,
                        request.getSubject(),
                        request.getBody(),
                        request.getFileName(),
                        request.getExcelFilePath(),
                        request.getSheetName(),
                        request.getEmailColumn()
                );

                updateStatus(row, statusCol, timeCol, "QUEUED", true);
            }

            // Write to temp file
            File tempFile = new File(excelFile.getParent(), "temp_" + excelFile.getName());
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                workbook.write(fos);
            }

            // Replace original file safely
            Path source = tempFile.toPath();
            Path target = excelFile.toPath();
            Files.copy(source, target, StandardCopyOption.REPLACE_EXISTING);
            Files.deleteIfExists(tempFile.toPath());

        } catch (Exception e) {
            throw new RuntimeException("Failed to process Excel file: " + e.getMessage(), e);
        }
    }

    private void updateStatus(Row row, int statusCol, int timeCol, String status, boolean updateTime) {
        row.createCell(statusCol).setCellValue(status);
        if (updateTime) {
            row.createCell(timeCol).setCellValue(
                    java.time.LocalDateTime.now().format(
                            java.time.format.DateTimeFormatter.ofPattern("dd-MM-yyyy hh.mm a"))
            );
        }
    }

    private int getColumnIndex(Row headerRow, String columnName) {
        for (Cell c : headerRow) {
            if (c.getCellType() == CellType.STRING &&
                    c.getStringCellValue().trim().equalsIgnoreCase(columnName.trim())) {
                return c.getColumnIndex();
            }
        }
        throw new RuntimeException("Column not found: " + columnName);
    }

    private int getOrCreateColumn(Row headerRow, String columnName) {
        for (Cell c : headerRow) {
            if (c.getCellType() == CellType.STRING &&
                    c.getStringCellValue().trim().equalsIgnoreCase(columnName.trim())) {
                return c.getColumnIndex();
            }
        }
        int idx = headerRow.getLastCellNum();
        if (idx < 0) idx = 0;
        headerRow.createCell(idx).setCellValue(columnName);
        return idx;
    }
}

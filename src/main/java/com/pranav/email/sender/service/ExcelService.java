package com.pranav.email.sender.service;

import com.pranav.email.sender.config.ExcelConstants;
import com.pranav.email.sender.dto.MailRequest;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.FileOutputStream;

@Service
public class ExcelService {

    private final EmailService emailService;

    public ExcelService(EmailService emailService) {
        this.emailService = emailService;
    }

    public void sendMailsFromExcel(MailRequest request) {

        try (Workbook workbook =
                     new XSSFWorkbook(new FileInputStream(ExcelConstants.FILE_PATH))) {

            Sheet sheet = workbook.getSheet(ExcelConstants.SHEET_NAME);
            Row headerRow = sheet.getRow(1);

            int emailCol = getColumnIndex(headerRow, ExcelConstants.EMAIL_COLUMN);
            int statusCol = getOrCreateColumn(headerRow, ExcelConstants.STATUS);
            int timeCol = getOrCreateColumn(headerRow, ExcelConstants.MAIL_SENT_TIME);

            for (int i = 2; i <= sheet.getLastRowNum(); i++) {

                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell emailCell = row.getCell(emailCol);
                if (emailCell == null) {
                    updateStatus(row, statusCol, timeCol, "INVALID", false);
                    continue;
                }

                String email = emailCell.getStringCellValue().trim();
                if (email.isEmpty() || !email.contains("@")) {
                    updateStatus(row, statusCol, timeCol, "INVALID", false);
                    continue;
                }

                emailService.sendMail(
                        email,
                        request.getSubject(),
                        request.getBody(),
                        request.getFileName()
                );

              
                updateStatus(row, statusCol, timeCol, "QUEUED", true);
            }

            try (FileOutputStream fos =
                         new FileOutputStream(ExcelConstants.FILE_PATH)) {
                workbook.write(fos);
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void updateStatus(Row row, int statusCol, int timeCol,
                              String status, boolean updateTime) {

        row.createCell(statusCol).setCellValue(status);

        if (!updateTime) return;

        row.createCell(timeCol)
                .setCellValue(java.time.LocalDateTime.now()
                        .format(java.time.format.DateTimeFormatter
                                .ofPattern("dd-MM-yyyy hh.mm a")));
    }

    private int getColumnIndex(Row headerRow, String columnName) {
        for (Cell c : headerRow) {
            if (c.getCellType() == CellType.STRING &&
                c.getStringCellValue().trim().equalsIgnoreCase(columnName)) {
                return c.getColumnIndex();
            }
        }
        throw new RuntimeException("Column not found");
    }

    private int getOrCreateColumn(Row headerRow, String columnName) {
        for (Cell c : headerRow) {
            if (c.getCellType() == CellType.STRING &&
                c.getStringCellValue().trim().equalsIgnoreCase(columnName)) {
                return c.getColumnIndex();
            }
        }
        int idx = headerRow.getLastCellNum();
        headerRow.createCell(idx).setCellValue(columnName);
        return idx;
    }
}

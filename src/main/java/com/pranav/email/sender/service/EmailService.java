package com.pranav.email.sender.service;

import jakarta.mail.internet.MimeMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.scheduling.annotation.Async;
import org.springframework.stereotype.Service;

import java.io.File;

@Service
public class EmailService {

    private final JavaMailSender mailSender;
    private final ExcelUpdateService excelUpdateService;

    public EmailService(JavaMailSender mailSender, ExcelUpdateService excelUpdateService) {
        this.mailSender = mailSender;
        this.excelUpdateService = excelUpdateService;
    }

    @Async("mailExecutor")
    public void sendMail(String to, String subject, String body, String filePath,
                         String excelPath, String sheetName, String emailColumn) {

        try {
            MimeMessage message = mailSender.createMimeMessage();
            MimeMessageHelper helper = new MimeMessageHelper(message, true);

            helper.setTo(to);
            helper.setSubject(subject);
            helper.setText(body);

            if (filePath != null && !filePath.isEmpty()) {
                File file = new File(filePath);
                if (file.exists()) {
                    helper.addAttachment(file.getName(), file);
                }
            }

            mailSender.send(message);

            excelUpdateService.updateStatus(excelPath, sheetName, emailColumn, to, "SENT");

        } catch (Exception e) {
            excelUpdateService.updateStatus(excelPath, sheetName, emailColumn, to, "FAILED");
        }
    }
}

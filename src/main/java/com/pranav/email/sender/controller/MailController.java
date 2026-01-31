package com.pranav.email.sender.controller;

import com.pranav.email.sender.dto.MailRequest;
import com.pranav.email.sender.service.ExcelService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/mail")
public class MailController {

    @Autowired
    private ExcelService excelService;

    @PostMapping("/send-from-excel")
    public String sendFromExcel(@RequestBody MailRequest request) {

        try {
            excelService.sendMailsFromExcel(request);
            return "Emails sent successfully and Excel updated";
        } catch (Exception e) {
            return "Error: " + e.getMessage();
        }
    }
}

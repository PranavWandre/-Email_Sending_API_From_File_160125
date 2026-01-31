package com.pranav.email.sender.dto;

public class MailRequest {

    private String subject;
    private String body;
    private String fileName;

    private String excelFilePath;
    private String sheetName;
    private String emailColumn;

    public String getSubject() { return subject; }
    public void setSubject(String subject) { this.subject = subject; }

    public String getBody() { return body; }
    public void setBody(String body) { this.body = body; }

    public String getFileName() { return fileName; }
    public void setFileName(String fileName) { this.fileName = fileName; }

    public String getExcelFilePath() { return excelFilePath; }
    public void setExcelFilePath(String excelFilePath) { this.excelFilePath = excelFilePath; }

    public String getSheetName() { return sheetName; }
    public void setSheetName(String sheetName) { this.sheetName = sheetName; }

    public String getEmailColumn() { return emailColumn; }
    public void setEmailColumn(String emailColumn) { this.emailColumn = emailColumn; }
}

package com.newlight77.kata.survey.utils;


import com.newlight77.kata.survey.model.AddressStatus;
import com.newlight77.kata.survey.model.Campaign;
import com.newlight77.kata.survey.model.Survey;
import com.newlight77.kata.survey.service.MailService;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

@Service
public class SaveToExcelFileService {

    private final MailService mailService;
    private static final DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");


    public SaveToExcelFileService(final MailService mailService){
        this.mailService = mailService;
    }

    //Method to write on file and send via mail
    public void writeFileAndSend(final Survey survey, final Workbook workbook){
        try {
            final File resultFile = new File(System.getProperty("java.io.tmpdir"), "survey-" + survey.getId() + "-" + dateTimeFormatter.format(LocalDate.now()) + ".xlsx");
            final FileOutputStream outputStream = new FileOutputStream(resultFile);
            workbook.write(outputStream);

            mailService.send(resultFile);
            resultFile.deleteOnExit();
        } catch(final Exception ex) {
            throw new RuntimeException("Errorr while trying to send email", ex);
        } finally {
            try {
                workbook.close();
            } catch(final Exception e) {
                // CANT HAPPEN
            }
        }
    }

    //create workbook
    public Workbook createWorkbook() {
        return new XSSFWorkbook();
    }

    //Create colomns for the survey
    public void setColumnWidths(Sheet sheet) {
        sheet.setColumnWidth(0, 10500);
        for (int i = 1; i <= 18; i++) {
            sheet.setColumnWidth(i, 6000);
        }
    }

    //Create style to apply on header
    public CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont font = ((XSSFWorkbook) workbook).createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 14);
        font.setBold(true);
        headerStyle.setFont(font);
        headerStyle.setWrapText(false);

        return headerStyle;
    }

    //Create style to apply on cells
    public CellStyle createTitleStyle(Workbook workbook) {
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        XSSFFont titleFont = ((XSSFWorkbook) workbook).createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 12);
        titleFont.setUnderline(FontUnderline.SINGLE);
        titleStyle.setFont(titleFont);

        return titleStyle;
    }

    //Create client section
    public void createClientSection(Sheet sheet, Survey survey, CellStyle titleStyle, CellStyle style) {
        Row row = sheet.createRow(2);
        Cell cell = row.createCell(0);
        cell.setCellValue("Client");
        cell.setCellStyle(titleStyle);

        Row clientRow = sheet.createRow(3);
        Cell nomClientRowLabel = clientRow.createCell(0);
        nomClientRowLabel.setCellValue(survey.getClient());
        nomClientRowLabel.setCellStyle(style);

        String clientAddress = survey.getClientAddress().getStreetNumber() + " "
                + survey.getClientAddress().getStreetName() + survey.getClientAddress().getPostalCode() + " "
                + survey.getClientAddress().getCity();

        Row clientAddressLabelRow = sheet.createRow(4);
        Cell clientAddressCell = clientAddressLabelRow.createCell(0);
        clientAddressCell.setCellValue(clientAddress);
        clientAddressCell.setCellStyle(style);
    }

    //Create Survey info cells and rows
    public void createSurveyInfo(Sheet sheet, Campaign campaign, CellStyle style) {
        Row row = sheet.createRow(6);
        Cell cell = row.createCell(0);
        cell.setCellValue("Number of surveys");
        cell = row.createCell(1);
        cell.setCellValue(campaign.getAddressStatuses().size());
    }

    //Create Survey Header values and apply style
    public void createSurveyHeader(Sheet sheet, CellStyle style) {
        Row surveyLabelRow = sheet.createRow(8);
        String[] headers = {"N° street", "streee", "Postal code", "City", "Status"};
        for (int i = 0; i < headers.length; i++) {
            Cell surveyLabelCell = surveyLabelRow.createCell(i);
            surveyLabelCell.setCellValue(headers[i]);
            surveyLabelCell.setCellStyle(style);
        }
    }

    //fill survey data  on corresponding row
    public void fillSurveyData(Sheet sheet, Campaign campaign, CellStyle style) {
        int startIndex = 9;
        int currentIndex = 0;

        for (AddressStatus addressStatus : campaign.getAddressStatuses()) {
            Row surveyRow = sheet.createRow(startIndex + currentIndex);
            surveyRow.createCell(0).setCellValue(addressStatus.getAddress().getStreetNumber());
            surveyRow.createCell(1).setCellValue(addressStatus.getAddress().getStreetName());
            surveyRow.createCell(2).setCellValue(addressStatus.getAddress().getPostalCode());
            surveyRow.createCell(3).setCellValue(addressStatus.getAddress().getCity());
            surveyRow.createCell(4).setCellValue(addressStatus.getStatus().toString());
            currentIndex++;
        }
    }

}